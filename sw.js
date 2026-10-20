// 판매가 검색기 서비스워커 — 앱 화면을 폰에 저장해 오프라인에서도 실행
// 버전을 올리면 이전 캐시가 지워지고 새 화면이 적용됩니다.
const CACHE = "rps-v1.0.38";
const SHELL = ["./", "index.html", "xlsx.full.min.js", "jszip.min.js", "qrcode.min.js", "manifest.json", "icon-192.png", "icon-512.png"];
const NAV_TIMEOUT_MS = 3000;   // 앱을 열 때 3초 안에 응답이 없으면 폰에 저장된 화면으로 바로 열기

self.addEventListener("install", e => {
  // 브라우저에 남은 옛 파일을 쓰지 않고 반드시 인터넷에서 새로 받아 저장 (업데이트 후 옛 화면이 다시 뜨는 현상 방지)
  e.waitUntil(caches.open(CACHE)
    .then(c => c.addAll(SHELL.map(u => new Request(u, { cache: "reload" }))))
    .then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// 네트워크에서 받아 정상(200)일 때만 저장
// fresh=true: 브라우저에 남은 옛 파일을 쓰지 않고 서버에 확인 (바뀐 게 없으면 304로 가볍게 끝남)
function fetchAndStore(req, fresh){
  const go = fresh
    ? fetch(req.url, { cache: "no-cache", credentials: "same-origin" }).then(r => r.redirected ? fetch(req) : r)
    : fetch(req);
  return go.then(r => {
    if (r && r.ok){ const cp = r.clone(); caches.open(CACHE).then(c => c.put(req, cp)); }
    return r;
  });
}

self.addEventListener("fetch", e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== "GET" || url.origin !== location.origin) return;   // 드라이브 요청은 그대로 통과
  if (url.pathname.endsWith("/sw.js")) return;                            // 새 버전 확인은 항상 네트워크

  const isPage = req.mode === "navigate" || url.pathname.endsWith(".html") || url.pathname.endsWith("/");
  if (!isPage){
    // 라이브러리, 아이콘: 캐시 우선
    e.respondWith(caches.match(req).then(r => r || fetch(req)));
    return;
  }
  const cached = () => caches.match(req).then(r => r || caches.match("index.html"));

  if (req.mode !== "navigate"){
    // 앱 안에서의 화면 파일 요청(업데이트 확인 등): 네트워크 우선, 실패 시 캐시
    e.respondWith(fetchAndStore(req).catch(cached));
    return;
  }
  // 앱 열기: 네트워크 우선, 단 3초 넘게 걸리면 저장된 화면으로 (네트워크 결과는 다음을 위해 계속 저장)
  e.respondWith(new Promise(resolve => {
    let done = false;
    const finish = r => { if (!done && r){ done = true; resolve(r); } };
    const net = fetchAndStore(req, true);
    const timer = setTimeout(() => { cached().then(finish); }, NAV_TIMEOUT_MS);
    net.then(r => { clearTimeout(timer); finish(r); })
       .catch(() => { clearTimeout(timer); cached().then(r => finish(r || Response.error())); });
    e.waitUntil(net.catch(() => {}));
  }));
});
