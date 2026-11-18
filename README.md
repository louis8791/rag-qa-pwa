# 📚 RAG 文件問答(PWA + RWD)

上傳 PDF / DOCX,選擇檢索策略,直接對文件提問。
使用 **Groq**(LLM)+ **ChromaDB**(向量庫)+ **Gradio**(介面)+ **FastAPI**(PWA 服務),
支援 **PWA**(可安裝成手機 / 電腦 App)與 **RWD**(手機、平板、桌機自動適應)。

## ✨ 功能

- 8 種檢索策略:混合、語意、關鍵字、重新排序、多查詢、壓縮、父子文檔、HyDE
- PDF(依頁)與 DOCX(依標題、含表格)智慧切片
- 回答附 `[片段n]` 引用,並可查看檢索到的原文片段
- 匯出問答紀錄為 Word / CSV
- 每位使用者各自獨立的文件與對話,多人同時使用不會互相干擾

## 📁 專案結構

```
.
├── app.py                  # Gradio + FastAPI 主程式
├── requirements.txt        # 完整模式套件
├── requirements-lite.txt   # 輕量模式套件(免費方案用)
├── render.yaml             # Render 部署設定
├── .gitignore
├── README.md
└── static/
    ├── manifest.webmanifest
    ├── sw.js
    ├── offline.html
    └── icons/ (192、512、maskable、apple-touch-icon、favicon)
```

## 🚀 部署到 Render

### ⚠️ 先看記憶體

| 模式 | 檢索 | 記憶體需求 | 建議方案 |
|---|---|---|---|
| **完整模式**(預設) | 語意 + 關鍵字 | 約 2GB(PyTorch + 嵌入模型) | standard 以上(請自行確認 Render 當前方案規格與價格) |
| **輕量模式** | 僅關鍵字 | < 512MB | free 即可 |

免費方案只有 512MB,跑完整模式會因記憶體不足而被強制重啟。

### 完整模式

1. 把整個資料夾(**含 `static/`**)推上 GitHub。
2. 在 Render 從 `louis8791/rag-qa-pwa` 建立新的 Blueprint，使用根目錄的 `render.yaml`。手動建立 Web Service 時，Root Directory 留空。
3. 公開服務讓訪客在介面使用自己的 [Groq 金鑰](https://console.groq.com/keys)。
4. 部署完成後,用 Render 網址開啟。

手動建立 Web Service 時（Root Directory 留空）:

| 欄位 | 值 |
|---|---|
| Build Command | `pip install -r requirements.txt && python -c "from sentence_transformers import SentenceTransformer as S; S('intfloat/multilingual-e5-small')"` |
| Start Command | `uvicorn app:app --host 0.0.0.0 --port $PORT` |
| Health Check Path | `/healthz` |
| 環境變數 | `PYTHON_VERSION=3.11.9`、`HF_HOME=/opt/render/project/src/.hf_cache` |

Build Command 會在建置階段先下載嵌入模型,啟動時就不用再下載。

### 輕量模式(免費方案)

- Build Command:`pip install -r requirements-lite.txt`
- Start Command:`uvicorn app:app --host 0.0.0.0 --port $PORT`
- 環境變數:`PYTHON_VERSION=3.11.9`、`EMBED_MODEL=none`

輕量模式下,「語意搜尋」「混合」「HyDE」等策略會退化為關鍵字檢索。

## 💻 本機執行

```bash
pip install -r requirements.txt
export GROQ_API_KEY=gsk_xxx        # Windows: set GROQ_API_KEY=gsk_xxx
python app.py                      # 開啟 http://localhost:7860
```

## ⚙️ 環境變數

| 變數 | 說明 | 預設 |
|---|---|---|
| `GROQ_API_KEY` | 伺服器預設的 Groq Key(使用者也可在介面自行輸入) | 無 |
| `EMBED_MODEL` | 嵌入模型;設為 `none` 啟用輕量模式 | `intfloat/multilingual-e5-small` |
| `CHROMA_PATH` | ChromaDB 儲存位置 | `./chroma_db` |
| `HF_HOME` | Hugging Face 模型快取位置 | `./.hf_cache` |
| `PORT` | 服務埠號(Render 自動提供) | `7860` |

換成 `intfloat/multilingual-e5-base` 或 `BAAI/bge-m3` 效果更好,但記憶體需求更高;
同時要把 `render.yaml` 的預下載指令改成同一個模型名稱。

## 🔐 安全與費用提醒

- 若設定了 `GROQ_API_KEY`,**所有訪客都會共用你的額度**。網站公開時,建議不要設定,改讓使用者在介面輸入自己的 Key。
- 伺服器的 Key **不會**傳到瀏覽器。使用者自己輸入的 Key 只存在他自己的連線 Session 中。
- 文件內容會送到 Groq 進行回答生成,請勿上傳機密文件。
- Render 的檔案系統是暫存的,重新部署或重啟後向量資料會清空(使用者需重新上傳文件)。

## 📱 PWA / RWD 說明

| 項目 | 實作 |
|---|---|
| Manifest | `static/manifest.webmanifest`(名稱、圖示、主題色、standalone) |
| Service Worker | `static/sw.js`,由 `/sw.js` 根路徑提供;斷線時顯示 `offline.html` |
| 安裝按鈕 | 符合安裝條件時,右上角出現「📲 安裝 App」 |
| RWD | ≤900px 左右欄改上下堆疊;≤640px 按鈕與輸入列加大、橫排元件改直排;支援瀏海安全區域 |

- 文件問答需要連網,離線時只會顯示離線頁。
- PWA 需要 HTTPS 才能安裝(Render 預設提供,本機 `localhost` 也可)。
- iPhone 請用 Safari:分享 → 加入主畫面。
- 修改 `sw.js` 後,把裡面的 `VERSION` 加 1(`v1` → `v2`),使用者端舊快取會自動更新。

## ⚠️ 其他注意事項

- 掃描版 PDF 沒有文字層,需先做 OCR。
- Render 免費方案閒置約 15 分鐘會休眠,再次開啟需要約 30~60 秒喚醒。
- 完整模式首次啟動時,嵌入模型在背景載入;若此時就上傳文件,畫面會顯示「嵌入模型載入中」,稍候即可。
