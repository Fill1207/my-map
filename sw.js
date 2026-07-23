/* 채맵 서비스워커 — 가볍게. 네트워크 우선(온라인이면 항상 최신), 실패 시 캐시(오프라인 대비).
   서버(script.google.com)·지오코딩·카카오 등 외부 요청은 건드리지 않고 그대로 네트워크로 보냄. */
const CACHE = "chaemap-v1";
const SHELL = ["./", "index.html", "koreamap.js", "places.enc.js",
  "manifest.webmanifest", "icon-180.png", "icon-192.png", "icon-512.png"];

self.addEventListener("install", e=>{
  e.waitUntil(caches.open(CACHE).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()));
});
self.addEventListener("activate", e=>{
  e.waitUntil(caches.keys()
    .then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k))))
    .then(()=>self.clients.claim()));
});
self.addEventListener("fetch", e=>{
  const req = e.request;
  if(req.method !== "GET") return;
  const url = new URL(req.url);
  if(url.origin !== location.origin) return; // 외부(서버·지도·지오코딩)는 그대로
  e.respondWith(
    fetch(req).then(res=>{
      const copy = res.clone();
      caches.open(CACHE).then(c=>c.put(req, copy));
      return res;
    }).catch(()=>caches.match(req).then(r=>r || caches.match("index.html")))
  );
});
