// 업무일지 앱: 화면 파일을 기기에 담아 두어 인터넷이 없어도 열리게 합니다.
// 화면을 고치면 VERSION 숫자를 올려 주세요.
const VERSION = "worklog-v1";
const SHELL = ["./", "index.html", "manifest.webmanifest", "icons/icon-192.png", "icons/icon-512.png", "icons/apple-touch-icon.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  // 구글 시트 저장소 요청은 항상 인터넷으로
  if (url.hostname.endsWith("google.com") || url.hostname.endsWith("googleusercontent.com")) return;

  // 앱 화면: 인터넷 먼저, 안 되면 담아 둔 것
  if (req.mode === "navigate" || url.origin === location.origin) {
    e.respondWith(
      fetch(req).then(res => {
        if (res.ok) { const copy = res.clone(); caches.open(VERSION).then(c => c.put(req, copy)); }
        return res;
      }).catch(() => caches.match(req).then(r => r || caches.match("index.html")))
    );
    return;
  }

  // 글꼴 등: 담아 둔 것 먼저
  if (url.hostname.endsWith("gstatic.com") || url.hostname.endsWith("googleapis.com")) {
    e.respondWith(
      caches.match(req).then(r => r || fetch(req).then(res => {
        const copy = res.clone(); caches.open(VERSION).then(c => c.put(req, copy)); return res;
      }))
    );
  }
});
