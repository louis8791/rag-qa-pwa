/* RAG 文件問答 - Service Worker
 *  1. 預先快取離線頁與圖示
 *  2. 頁面導覽:網路優先,斷線時顯示 offline.html
 *  3. 靜態資源(Gradio 前端檔案、圖示):快取優先 + 背景更新
 *  4. API / 上傳 / 即時資料(/gradio_api、/queue 等)一律不快取、不介入
 * 更新時請把 VERSION 加 1,舊快取會自動清除。
 */
const VERSION = "v1";
const CACHE = "rag-pwa-" + VERSION;
const PRECACHE = [
  "/offline.html",
  "/manifest.webmanifest",
  "/pwa-static/icons/icon-192.png",
  "/pwa-static/icons/icon-512.png",
  "/pwa-static/icons/apple-touch-icon.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE).then((cache) => cache.addAll(PRECACHE)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k.startsWith("rag-pwa-") && k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

function isStaticAsset(url) {
  return (
    url.pathname.startsWith("/pwa-static/") ||
    url.pathname.startsWith("/assets/") ||
    url.pathname.startsWith("/_app/") ||
    url.pathname.startsWith("/static/") ||
    url.pathname === "/manifest.webmanifest"
  );
}

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;

  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  if (url.pathname.startsWith("/gradio_api") || url.pathname.startsWith("/queue") || url.pathname.startsWith("/run")) return;

  if (req.mode === "navigate") {
    event.respondWith(fetch(req).catch(() => caches.match("/offline.html")));
    return;
  }

  if (isStaticAsset(url)) {
    event.respondWith(
      caches.open(CACHE).then((cache) =>
        cache.match(req).then((cached) => {
          const network = fetch(req)
            .then((res) => {
              if (res && res.ok) cache.put(req, res.clone());
              return res;
            })
            .catch(() => cached);
          return cached || network;
        })
      )
    );
  }
});
