// 판매가 검색기 서비스워커 — 앱 화면을 폰에 저장해 오프라인에서도 실행
// 버전을 올리면 이전 캐시가 지워지고 새 화면이 적용됩니다.
const CACHE = "rps-v1.0.4";
const SHELL = ["./", "index.html", "xlsx.full.min.js", "jszip.min.js", "manifest.json", "icon-192.png", "icon-512.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;  // 드라이브 요청은 그대로 통과
  // 화면(html)은 네트워크 우선 → 업데이트 즉시 반영, 실패 시 캐시
  if (e.request.mode === "navigate" || url.pathname.endsWith(".html")) {
    e.respondWith(
      fetch(e.request).then(r => { const cp = r.clone(); caches.open(CACHE).then(c => c.put(e.request, cp)); return r; })
        .catch(() => caches.match(e.request).then(r => r || caches.match("index.html")))
    );
    return;
  }
  // 나머지(라이브러리, 아이콘)는 캐시 우선
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
});
