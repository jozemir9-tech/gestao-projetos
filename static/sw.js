// Service worker do PWA (Gestão de Projetos).
// Estratégia: network-first com cache de reserva — conteúdo sempre fresco
// quando há rede; offline cai para o que já foi visto.
const CACHE = "gestao-projetos-v1";

self.addEventListener("install", (evento) => {
  self.skipWaiting();
});

self.addEventListener("activate", (evento) => {
  evento.waitUntil(
    caches
      .keys()
      .then((chaves) =>
        Promise.all(chaves.filter((c) => c !== CACHE).map((c) => caches.delete(c)))
      )
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (evento) => {
  const req = evento.request;
  if (req.method !== "GET" || !req.url.startsWith("http")) return;
  // Dados do Streamlit vêm por WebSocket — não interceptar.
  if (req.headers.get("accept")?.includes("text/event-stream")) return;

  evento.respondWith(
    fetch(req)
      .then((resp) => {
        const copia = resp.clone();
        caches
          .open(CACHE)
          .then((cache) => cache.put(req, copia))
          .catch(() => {});
        return resp;
      })
      .catch(() =>
        caches.match(req).then((salvo) => salvo || Response.error())
      )
  );
});
