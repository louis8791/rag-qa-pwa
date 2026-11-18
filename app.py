"""多策略 RAG 文件問答系統 (Groq + ChromaDB + Gradio) —— PWA / RWD 版,可部署到 Render

設計重點(相較 Colab 單一儲存格版):
  1. 每位使用者各自擁有獨立的文件與狀態(Session),多人同時使用不會互相覆蓋。
  2. 嵌入模型改為背景執行緒載入,服務可立即啟動並通過 Render 健康檢查。
  3. Groq Key 不會預填到瀏覽器(避免外洩);使用者可自行輸入,否則使用伺服器環境變數。
  4. 以 FastAPI 提供 PWA 檔案(manifest / sw.js / 離線頁),再把 Gradio 掛在根路徑。
  5. EMBED_MODEL 設為 none 即為「輕量模式」(僅關鍵字檢索,不需 torch,可跑 512MB 記憶體)。
"""
import csv
import os
import re
import tempfile
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
os.environ.setdefault("HF_HOME", str(BASE_DIR / ".hf_cache"))
os.environ.setdefault("ANONYMIZED_TELEMETRY", "False")
os.environ.setdefault("TOKENIZERS_PARALLELISM", "false")

import numpy as np
import gradio as gr
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph
from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from groq import Groq
from pypdf import PdfReader
from sklearn.feature_extraction.text import TfidfVectorizer

# ====================== 設定 ======================
EMBED_MODEL = os.getenv("EMBED_MODEL", "intfloat/multilingual-e5-small").strip()  # 可換 multilingual-e5-base、BAAI/bge-m3
LITE = EMBED_MODEL.lower() in ("", "none", "lite", "tfidf")                       # 輕量模式:不載入嵌入模型
MODELS = ["openai/gpt-oss-20b", "openai/gpt-oss-120b", "qwen/qwen3.8-27b"]
CHUNK, OVERLAP, CHILD, CHILD_OVERLAP, MAX_CTX = 500, 100, 200, 40, 7000
CHROMA_PATH = os.getenv("CHROMA_PATH", str(BASE_DIR / "chroma_db"))
GROQ_KEY = os.getenv("GROQ_API_KEY", "").strip()

STRATEGIES = {
    "hybrid": ("⚡ 混合搜尋", "語意 + 關鍵字,RRF 融合排名(推薦)"),
    "semantic": ("🔍 語意搜尋", "向量相似度檢索"),
    "tfidf": ("📊 關鍵字", "字元 n-gram,適合專有名詞、條號、數字"),
    "rerank": ("🎯 重新排序", "LLM 對候選片段批次評分後重排"),
    "multi_query": ("🔄 多查詢", "LLM 改寫 3 個角度問題聯合檢索"),
    "compress": ("✂️ 壓縮", "LLM 抽取每段最相關的句子"),
    "parent_child": ("📂 父子文檔", "小片段定位,回傳完整大片段"),
    "hyde": ("💡 HyDE", "先生成假設答案再檢索"),
}
_SENT = re.compile(r"(?<=[。！？!?；;])|\n+")


def split_text(text, size, overlap):
    sents = [s.strip() for s in _SENT.split(re.sub(r"[ \t　\xa0]+", " ", text)) if s and s.strip()]
    chunks, buf, fresh = [], "", 0
    for p in (s[i:i + size] for s in sents for i in range(0, len(s), size)):
        if fresh and len(buf) + len(p) + 1 > size:
            chunks.append(buf)
            buf, fresh = (buf[-overlap:].lstrip() if overlap else ""), 0
        buf = f"{buf} {p}" if buf else p
        fresh += len(p)
    return chunks + ([buf.strip()] if fresh else [])


# ====================== 文件讀取 ======================
def read_pdf(path):
    return [(f"第 {i} 頁", t) for i, p in enumerate(PdfReader(str(path)).pages, 1)
            if (t := p.extract_text() or "").strip()]


def read_docx(path):
    doc, segs, head, buf, nt = Document(str(path)), [], "前言", [], 0

    def flush():
        if buf:
            segs.append((head, "\n".join(buf)))
            buf.clear()

    for el in doc.element.body.iterchildren():
        if isinstance(el, CT_P):
            para = Paragraph(el, doc)
            txt = para.text.strip()
            if not txt:
                continue
            style = (para.style.name if para.style is not None else "") or ""
            if style.lower().startswith(("heading", "title")) or style.startswith(("標題", "标题")):
                flush()
                head = txt[:40]
            buf.append(txt)
        elif isinstance(el, CT_Tbl):
            nt += 1
            rows = []
            for row in Table(el, doc).rows:
                cells = []
                for c in row.cells:
                    t = re.sub(r"\s+", " ", c.text).strip()
                    if t and (not cells or cells[-1] != t):
                        cells.append(t)
                if cells:
                    rows.append(" | ".join(cells))
            if rows:
                flush()
                segs.append((f"表格 {nt}（{head}）", "\n".join(rows)))
    flush()
    return segs


# ====================== 共用引擎:嵌入模型 + ChromaDB ======================
class Engine:
    """全站共用一份(模型很吃記憶體),在背景執行緒載入。"""

    def __init__(self):
        self.ready, self.error = threading.Event(), None
        self.embedder = self.col = self.ccol = None
        self.e5, self.lock = "e5" in EMBED_MODEL.lower(), threading.Lock()
        threading.Thread(target=self._boot, daemon=True).start()

    def _boot(self):
        try:
            if not LITE:
                import chromadb
                import torch
                from sentence_transformers import SentenceTransformer
                device = "cuda" if torch.cuda.is_available() else "cpu"
                self.embedder = SentenceTransformer(EMBED_MODEL, device=device)
                self.embedder.max_seq_length = min(512, self.embedder.max_seq_length or 512)
                db = chromadb.PersistentClient(path=CHROMA_PATH)
                slug = re.sub(r"[^a-zA-Z0-9]+", "_", EMBED_MODEL.split("/")[-1])[:40]
                self.col = db.get_or_create_collection(f"p_{slug}", metadata={"hnsw:space": "cosine"})
                self.ccol = db.get_or_create_collection(f"c_{slug}", metadata={"hnsw:space": "cosine"})
        except Exception as e:  # noqa
            self.error = e
        finally:
            self.ready.set()

    def enc(self, texts, query=False):
        if self.e5:
            texts = [("query: " if query else "passage: ") + t for t in texts]
        with self.lock:
            return self.embedder.encode(texts, batch_size=64, normalize_embeddings=True,
                                        show_progress_bar=False).tolist()

    def add(self, col, ids, docs, metas, b=2000):
        for i in range(0, len(docs), b):
            col.add(ids=ids[i:i + b], documents=docs[i:i + b], metadatas=metas[i:i + b],
                    embeddings=self.enc(docs[i:i + b]))

    def query(self, sid, qs, k, total, child):
        res = (self.ccol if child else self.col).query(
            query_embeddings=self.enc(qs, True), n_results=max(1, min(k, total)),
            where={"sid": sid}, include=["metadatas"])
        return [[int(m["idx"]) for m in ms] for ms in res["metadatas"]]

    def drop(self, sid):
        if not sid or self.col is None:
            return
        for c in (self.col, self.ccol):
            try:
                c.delete(where={"sid": sid})
            except Exception:  # noqa
                pass


ENGINE = Engine()


def list_models(client):
    if not client:
        return []
    try:
        ids = [m.id for m in client.models.list().data if getattr(m, "active", True) is not False]
    except Exception:  # noqa
        return []
    return sorted(i for i in ids if not any(s in i.lower() for s in ("whisper", "tts", "orpheus", "guard", "playai")))


# ====================== 每位使用者獨立的 Session ======================
class Session:
    def __init__(self):
        self.client, self.model = None, MODELS[0]
        self.sid, self.name = None, ""
        self.chunks, self.labels, self.n_child = [], [], 0
        self.tfidf, self.tm = None, None
        self.set_key(GROQ_KEY)

    def close(self):
        ENGINE.drop(self.sid)
        self.sid = None

    # ── Groq ──
    def set_key(self, key):
        self.client = Groq(api_key=key.strip()) if key and key.strip() else None

    def llm(self, prompt, max_tokens=300, temp=0.3, system=None, raise_err=False):
        if not self.client:
            return None
        kw = dict(model=self.model, temperature=temp, max_completion_tokens=max_tokens,
                  messages=([{"role": "system", "content": system}] if system else [])
                  + [{"role": "user", "content": prompt}])
        if "gpt-oss" in self.model:  # 推理模型:推理 token 也計入上限
            kw.update(max_completion_tokens=max(max_tokens * 3, 1024), reasoning_effort="low", include_reasoning=False)
        try:
            try:
                r = self.client.chat.completions.create(**kw)
            except Exception as e:
                if "reasoning" not in str(e).lower():
                    raise
                kw.pop("reasoning_effort", None)
                kw.pop("include_reasoning", None)
                r = self.client.chat.completions.create(**kw)
            return re.sub(r"<think>.*?</think>", "", r.choices[0].message.content or "", flags=re.S).strip()
        except Exception:
            if raise_err:
                raise
            return None

    # ── 文件載入 ──
    def load(self, path):
        t0, path = time.time(), Path(path)
        ext = path.suffix.lower()
        if ext not in (".pdf", ".docx"):
            return "✗ 僅支援 PDF / DOCX"
        try:
            segs = read_pdf(path) if ext == ".pdf" else read_docx(path)
            pairs = [(c, lab) for lab, txt in segs for c in split_text(txt, CHUNK, OVERLAP)]
            if not pairs:
                return "✗ 無可擷取文字(可能是掃描檔,需先 OCR)"
            self.close()
            self.sid, self.name = f"{int(time.time())}_{uuid.uuid4().hex[:6]}", path.name
            self.chunks, self.labels = [p[0] for p in pairs], [p[1] for p in pairs]
            self.n_child = 0
            if not LITE:
                ENGINE.add(ENGINE.col, [f"{self.sid}_{i}" for i in range(len(pairs))], self.chunks,
                           [{"sid": self.sid, "idx": i} for i in range(len(pairs))])
                kids = [(f"{self.sid}_{p}_{c}", t, p) for p, par in enumerate(self.chunks)
                        for c, t in enumerate(split_text(par, CHILD, CHILD_OVERLAP))]
                self.n_child = len(kids)
                ENGINE.add(ENGINE.ccol, [k[0] for k in kids], [k[1] for k in kids],
                           [{"sid": self.sid, "idx": k[2]} for k in kids])
            self.tfidf = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4), sublinear_tf=True, max_features=50000)
            self.tm = self.tfidf.fit_transform(self.chunks)
            return f"✓ {self.name}\n{len(self.chunks)} 片段 · {time.time() - t0:.1f}s"
        except Exception as e:
            return f"✗ 載入失敗:{type(e).__name__}: {e}"

    # ── 檢索 ──
    def vec(self, qs, k, child=False):
        if LITE or ENGINE.embedder is None:  # 輕量模式:語意檢索退化為關鍵字檢索
            return [self.kw(q, k) for q in qs]
        total = self.n_child if child else len(self.chunks)
        if not total:
            return [[] for _ in qs]
        return ENGINE.query(self.sid, qs, k, total, child)

    def kw(self, q, k):
        s = (self.tm @ self.tfidf.transform([q]).T).toarray().ravel()
        return [int(i) for i in np.argsort(-s)[:k] if s[i] > 0]

    @staticmethod
    def rrf(lists, k):
        sc = {}
        for lst in lists:
            for r, i in enumerate(dict.fromkeys(lst)):
                sc[i] = sc.get(i, 0) + 1 / (61 + r)
        return sorted(sc, key=sc.get, reverse=True)[:k]

    def retrieve(self, q, strategy, k):  # 回傳 [(片段索引, 文字)]
        hybrid = lambda n: self.rrf([self.vec([q], n * 2)[0], self.kw(q, n * 2)], n)
        if strategy == "semantic":
            idx = self.vec([q], k)[0]
        elif strategy == "tfidf":
            idx = self.kw(q, k)
        elif strategy == "parent_child":
            idx = list(dict.fromkeys(self.vec([q], k * 4, True)[0]))[:k]
        elif strategy == "multi_query":
            r = self.llm(f"將以下問題改寫成 3 個角度不同的繁體中文問題,每行一題,不加編號:\n{q}", 200, 0.7) or ""
            extra = [e for e in (l.strip("-•*·0123456789.、)） ") for l in r.splitlines()) if len(e) > 3][:3]
            idx = self.rrf(self.vec([q] + extra, k * 2), k)
        elif strategy == "hyde":
            h = self.llm(f"請針對以下問題寫一段約 100 字、語氣像文件原文的假設性答案(繁體中文):\n{q}", 250, 0.7)
            idx = self.rrf(self.vec([q, h] if h else [q], k * 2), k)
        elif strategy == "rerank":
            cands = hybrid(min(max(k * 3, 6), 12))
            idx = cands[:k]
            if self.client and len(cands) > 1:
                lst = "\n\n".join(f"[{n}] {self.chunks[i][:400]}" for n, i in enumerate(cands, 1))
                r = self.llm(f"問題:{q}\n\n候選片段:\n{lst}\n\n評估每段與問題的相關度(0-10),只輸出「編號:分數」每行一個。",
                             15 * len(cands), 0) or ""
                sc = {int(a): float(b) for a, b in re.findall(r"\[?(\d+)\]?\s*[:：]\s*(\d+(?:\.\d+)?)", r)}
                idx = [cands[n] for n in sorted(range(len(cands)), key=lambda n: (-sc.get(n + 1, -1), n))[:k]]
        else:
            idx = hybrid(k)
        hits = [(i, self.chunks[i]) for i in idx]
        if strategy == "compress" and self.client and hits:
            def comp(h):
                r = self.llm(f"從以下文本逐字擷取與問題「{q}」最相關的 1-3 句,不改寫不解釋;完全無關請輸出「無」。\n\n{h[1]}", 250, 0)
                return (h[0], r if r and r != "無" else h[1][:300])
            with ThreadPoolExecutor(min(4, len(hits))) as ex:
                hits = list(ex.map(comp, hits))
        return hits

    def answer(self, q, strategy, k, model):
        if not self.chunks:
            return "請先上傳 PDF 或 DOCX 文件。", ""
        self.model, t0 = model or MODELS[0], time.time()
        hits = self.retrieve(q, strategy, int(k))
        t1 = time.time()
        blocks, used = [], 0
        for n, (i, txt) in enumerate(hits, 1):
            b = f"[片段{n}｜{self.labels[i]}]\n{txt}"
            if used and used + len(b) > MAX_CTX:
                break
            blocks.append(b)
            used += len(b)
        ctx = "\n\n".join(blocks)
        src = f"📄 {self.name} · {STRATEGIES[strategy][0]} · {len(blocks)} 片段 · 檢索 {t1 - t0:.2f}s\n{'─' * 40}\n\n{ctx}"
        if not hits:
            return "找不到相關片段,請換個問法或策略。", src
        if not self.client:
            return "⚠ 尚未設定 Groq API Key(檢索結果請看下方片段)", src
        try:
            ans = self.llm(f"請只根據上下文回答。若上下文沒有相關資訊,請說「文件中未提及」,不要編造。\n"
                           f"每個重點後以 [片段n] 標註依據,使用繁體中文條列整理。\n\n上下文:\n{ctx}\n\n問題:{q}",
                           1024, 0.2, "你是嚴謹的文件問答助手,只依據提供的上下文作答。", True)
        except Exception as e:
            return f"✗ 生成失敗:{type(e).__name__}: {e}", src
        return f"{ans or '(模型未回傳內容)'}\n\n<sub>{self.model} · 檢索 {t1 - t0:.1f}s · 生成 {time.time() - t1:.1f}s</sub>", src


# ====================== 匯出:Word / CSV ======================
def _out_path(ext):
    return os.path.join(tempfile.gettempdir(), f"RAG問答紀錄_{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:4]}.{ext}")


def export_csv(records):
    if not records:
        gr.Warning("目前沒有問答紀錄可匯出")
        return None
    path = _out_path("csv")
    with open(path, "w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig:Excel 開啟中文不亂碼
        w = csv.DictWriter(f, fieldnames=list(records[0]))
        w.writeheader()
        w.writerows(records)
    return path


def _md_runs(par, text):  # 將 **粗體** 轉成 Word 粗體
    for i, part in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if part:
            par.add_run(part).bold = bool(i % 2)


def export_docx(records, with_src=True):
    if not records:
        gr.Warning("目前沒有問答紀錄可匯出")
        return None
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Microsoft JhengHei", Pt(11)
    st.element.rPr.rFonts.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia", "Microsoft JhengHei")
    doc.add_heading("RAG 文件問答紀錄", 0)
    meta = doc.add_paragraph(f"文件:{records[-1]['文件']}　·　匯出時間:{datetime.now():%Y-%m-%d %H:%M}　·　共 {len(records)} 題")
    meta.runs[0].font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
    for n, r in enumerate(records, 1):
        doc.add_heading(f"Q{n}．{r['問題']}", level=1)
        info = doc.add_paragraph(f"{r['時間']}　·　{r['策略']}　·　{r['模型']}")
        info.runs[0].font.size, info.runs[0].font.color.rgb = Pt(9), RGBColor(0x6B, 0x72, 0x80)
        for line in r["回答"].splitlines():
            line = line.rstrip()
            if not line.strip():
                continue
            if m := re.match(r"^\s*#{1,6}\s+(.*)", line):
                doc.add_heading(m.group(1).replace("**", ""), level=2)
                continue
            if m := re.match(r"^(\s*)[-*•]\s+(.*)", line):
                par = doc.add_paragraph(style="List Bullet 2" if len(m.group(1)) >= 2 else "List Bullet")
                _md_runs(par, m.group(2))
            elif m := re.match(r"^\s*\d+[.)]\s+(.*)", line):
                par = doc.add_paragraph(style="List Number")
                _md_runs(par, m.group(1))
            else:
                _md_runs(doc.add_paragraph(), line)
        if with_src and r["參考片段"]:
            h = doc.add_paragraph()
            run = h.add_run("參考片段")
            run.bold, run.font.size = True, Pt(9)
            for blk in r["參考片段"].split("\n\n"):
                p = doc.add_paragraph(blk.strip())
                p.paragraph_format.left_indent = Pt(12)
                for run in p.runs:
                    run.font.size, run.font.color.rgb = Pt(8.5), RGBColor(0x55, 0x5B, 0x66)
    path = _out_path("docx")
    doc.save(path)
    return path


# ====================== PWA:<head> ======================
PWA_HEAD = """
<link rel="manifest" href="/manifest.webmanifest">
<meta name="theme-color" content="#059669">
<meta name="application-name" content="RAG問答">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="RAG問答">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<link rel="apple-touch-icon" href="/pwa-static/icons/apple-touch-icon.png">
<link rel="icon" type="image/png" href="/pwa-static/icons/favicon.png">
<style>
  #pwa-install-btn {
    position: fixed; right: 12px; top: calc(12px + env(safe-area-inset-top, 0px));
    z-index: 9999; min-height: 44px; padding: 0 16px; border: 0; border-radius: 999px;
    background: #059669; color: #fff; font-size: .95rem; font-weight: 600;
    box-shadow: 0 6px 20px rgba(0,0,0,.25); cursor: pointer;
  }
</style>
<script>
(function () {
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", function () {
      navigator.serviceWorker.register("/sw.js", { scope: "/" })
        .catch(function (e) { console.warn("Service Worker 註冊失敗:", e); });
    });
  }
  var deferredPrompt = null;
  function showInstallButton() {
    if (document.getElementById("pwa-install-btn")) return;
    var btn = document.createElement("button");
    btn.id = "pwa-install-btn";
    btn.type = "button";
    btn.textContent = "📲 安裝 App";
    btn.onclick = function () {
      btn.remove();
      if (!deferredPrompt) return;
      deferredPrompt.prompt();
      deferredPrompt.userChoice.finally(function () { deferredPrompt = null; });
    };
    document.body.appendChild(btn);
  }
  window.addEventListener("beforeinstallprompt", function (e) {
    e.preventDefault();
    deferredPrompt = e;
    if (document.body) showInstallButton();
    else document.addEventListener("DOMContentLoaded", showInstallButton);
  });
  window.addEventListener("appinstalled", function () {
    var b = document.getElementById("pwa-install-btn");
    if (b) b.remove();
  });
})();
</script>
"""

# ====================== RWD:CSS ======================
CSS = """
.gradio-container{
  max-width:1200px!important;margin:auto;
  padding-left:max(12px,env(safe-area-inset-left))!important;
  padding-right:max(12px,env(safe-area-inset-right))!important;
}
#hdr{padding:18px 4px 6px}
#hdr h1{font-size:26px;margin:0;font-weight:700}
#hdr p{color:#6b7280;margin:4px 0 0;font-size:14px}
.side{border-radius:14px!important}
#send{min-width:90px}
#chat{height:clamp(320px,62vh,640px)!important}
.main-row{flex-wrap:wrap!important}
footer{display:none!important}

/* 平板以下:左右欄改為上下堆疊 */
@media (max-width:900px){
  .side-col,.chat-col{min-width:100%!important;flex:1 1 100%!important}
  #hdr h1{font-size:22px}
}

/* 手機 */
@media (max-width:640px){
  #hdr{padding:10px 2px 2px;padding-right:110px}   /* 右側留位置給「安裝 App」按鈕 */
  #hdr h1{font-size:19px}
  #hdr p{font-size:12px}
  #chat{height:clamp(280px,55vh,520px)!important}
  .stack-mobile{flex-wrap:wrap!important}
  .stack-mobile>*{min-width:100%!important;flex:1 1 100%!important}
  button{min-height:44px}
  .gradio-container{padding-left:8px!important;padding-right:8px!important}
}
"""

THEME = gr.themes.Soft(primary_hue="emerald", neutral_hue="slate", radius_size="lg",
                       font=[gr.themes.GoogleFont("Noto Sans TC"), "sans-serif"])


# ====================== 事件處理 ======================
def on_key(key, cur, sess):
    sess = sess or Session()
    k = (key or "").strip()
    if not k:
        sess.set_key(GROQ_KEY)
        return ("✓ 使用伺服器預設 Key" if GROQ_KEY else "⚠ 尚未設定 Key"), gr.update(), sess
    sess.set_key(k)
    ms = list_models(sess.client)
    if not ms:
        return "✗ Key 無效或無法取得模型", gr.update(), sess
    return (f"✓ Key 有效 · {len(ms)} 個模型",
            gr.update(choices=ms, value=cur if cur in ms else (MODELS[0] if MODELS[0] in ms else ms[0])), sess)


def on_upload(f, sess):
    sess = sess or Session()
    if f is None:
        yield "⚠ 請選擇檔案", sess
        return
    if not ENGINE.ready.is_set():
        yield "⏳ 嵌入模型載入中(首次啟動較久),請稍候…", sess
    ENGINE.ready.wait()
    if ENGINE.error:
        yield f"✗ 嵌入模型載入失敗:{type(ENGINE.error).__name__}: {ENGINE.error}", sess
        return
    yield "⏳ 文件處理中…", sess
    yield sess.load(f if isinstance(f, str) else f.name), sess


def on_ask(q, history, records, strategy, k, model, sess):
    q = (q or "").strip()
    if not q:
        return "", history, gr.update(), records
    sess = sess or Session()
    ans, src = sess.answer(q, strategy, k, model)
    records = records + [{"時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "文件": sess.name, "問題": q,
                          "回答": re.sub(r"\n*<sub>.*?</sub>", "", ans, flags=re.S).strip(),
                          "策略": STRATEGIES[strategy][0], "模型": sess.model,
                          "參考片段": src.split("─" * 40, 1)[-1].strip()}]
    return "", history + [{"role": "user", "content": q}, {"role": "assistant", "content": ans}], src, records


init_models = list_models(Groq(api_key=GROQ_KEY)) if GROQ_KEY else []
init_models = init_models or MODELS
mode_note = "⚡ 輕量模式(僅關鍵字檢索)" if LITE else "Groq × ChromaDB"

# ====================== Gradio 介面 ======================
with gr.Blocks(title="RAG 文件問答", theme=THEME, css=CSS, head=PWA_HEAD) as demo:
    sess_state = gr.State(None, delete_callback=lambda s: s.close() if s else None)
    records = gr.State([])

    gr.HTML(f"<div id='hdr'><h1>📚 RAG 文件問答</h1>"
            f"<p>上傳 PDF / DOCX,選擇檢索策略,直接提問 · {mode_note}</p></div>")
    with gr.Row(equal_height=False, elem_classes="main-row"):
        with gr.Column(scale=1, min_width=300, elem_classes=["side", "side-col"]):
            with gr.Accordion("🔑 Groq API Key", open=not GROQ_KEY):
                key = gr.Textbox(type="password", show_label=False, placeholder="gsk_…(Enter 套用;留空則用伺服器預設)")
                key_status = gr.Markdown("✓ 使用伺服器預設 Key" if GROQ_KEY else "⚠ 尚未設定 · 請輸入你的 Groq API Key")
            model = gr.Dropdown(init_models, value=MODELS[0] if MODELS[0] in init_models else init_models[0],
                                allow_custom_value=True, label="模型")
            file = gr.File(label="上傳文件", file_types=[".pdf", ".docx"], type="filepath", height=120)
            status = gr.Textbox(show_label=False, interactive=False, lines=2, placeholder="尚未載入文件")
            strategy = gr.Radio([(v[0], k) for k, v in STRATEGIES.items()], value="hybrid", label="檢索策略")
            desc = gr.Markdown(f"<small>{STRATEGIES['hybrid'][1]}</small>")
            topk = gr.Slider(1, 10, 4, step=1, label="片段數量 Top-K")
        with gr.Column(scale=2, elem_classes="chat-col"):
            chat = gr.Chatbot(type="messages", height="60vh", show_label=False, elem_id="chat",
                              placeholder="**上傳文件後開始提問**<br>例如:這份文件的重點是什麼?")
            with gr.Row():
                q = gr.Textbox(show_label=False, placeholder="輸入問題,Enter 送出", scale=6, container=False)
                send = gr.Button("送出", variant="primary", scale=1, elem_id="send", min_width=80)
            with gr.Row(elem_classes="stack-mobile"):
                gr.Examples([["這份文件的重點是什麼?"], ["文件提到哪些規定或限制?"], ["有哪些風險或注意事項?"]], q, label=None)
                clear = gr.Button("🗑 清除對話", size="sm", variant="secondary")
            with gr.Accordion("🔎 檢索到的片段", open=False):
                src = gr.Textbox(show_label=False, lines=12, interactive=False)
            with gr.Row(elem_classes="stack-mobile"):
                with_src = gr.Checkbox(True, label="Word 包含參考片段", scale=2)
                btn_docx = gr.Button("📄 匯出 Word", size="sm", scale=1)
                btn_csv = gr.Button("📊 匯出 CSV", size="sm", scale=1)
            out_file = gr.File(label="下載檔案", visible=False)
            with gr.Accordion("📲 如何安裝成手機 / 電腦 App", open=False):
                gr.Markdown(
                    "- **Android(Chrome)**:點右上角「📲 安裝 App」,或選單 → 加到主畫面\n"
                    "- **iPhone / iPad(Safari)**:點「分享」→「加入主畫面」\n"
                    "- **電腦(Chrome / Edge)**:網址列右側的安裝圖示 → 安裝")

    # ---------------------- 事件綁定 ----------------------
    demo.load(lambda: Session(), None, sess_state)
    key.submit(on_key, [key, model, sess_state], [key_status, model, sess_state])
    key.blur(on_key, [key, model, sess_state], [key_status, model, sess_state])
    file.upload(on_upload, [file, sess_state], [status, sess_state])
    strategy.change(lambda s: f"<small>{STRATEGIES[s][1]}</small>", strategy, desc)
    for ev in (q.submit, send.click):
        ev(on_ask, [q, chat, records, strategy, topk, model, sess_state], [q, chat, src, records])
    clear.click(lambda: ([], "", [], gr.update(value=None, visible=False)), None, [chat, src, records, out_file])
    show = lambda p: gr.update(value=p, visible=bool(p))
    btn_docx.click(lambda r, w: show(export_docx(r, w)), [records, with_src], out_file)
    btn_csv.click(lambda r: show(export_csv(r)), records, out_file)

demo.queue(default_concurrency_limit=4)

# ====================== FastAPI:PWA 路由 + 掛載 Gradio ======================
# 注意:PWA 檔案路由必須在 mount_gradio_app(path="/") 之前宣告
app = FastAPI(title="RAG 文件問答")


@app.get("/healthz", include_in_schema=False)
def healthz():
    return {"status": "ok", "embedder_ready": ENGINE.ready.is_set(), "lite": LITE}


@app.get("/manifest.webmanifest", include_in_schema=False)
def manifest():
    return FileResponse(STATIC_DIR / "manifest.webmanifest", media_type="application/manifest+json")


@app.get("/sw.js", include_in_schema=False)
def service_worker():
    return FileResponse(STATIC_DIR / "sw.js", media_type="application/javascript",
                        headers={"Service-Worker-Allowed": "/", "Cache-Control": "no-cache"})


@app.get("/offline.html", include_in_schema=False)
def offline():
    return FileResponse(STATIC_DIR / "offline.html", media_type="text/html")


app.mount("/pwa-static", StaticFiles(directory=STATIC_DIR), name="pwa-static")
app = gr.mount_gradio_app(app, demo, path="/")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 7860)))
