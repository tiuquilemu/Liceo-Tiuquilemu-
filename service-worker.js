// Actualización automática del sistema de asistencia QR.
// Al activarse, elimina solo las cachés antiguas de esta plataforma y recarga
// las pestañas abiertas para que todos los computadores usen la misma versión.
const APP_VERSION = '19';
const CACHE_NAME = 'asistencia-qr-shell-v' + APP_VERSION;
const SHELL_FILES = [
  './',
  './index.html',
  './app.js?v=' + APP_VERSION,
  './manifest.webmanifest?v=' + APP_VERSION,
  './vendor/jsQR.js',
  './vendor/chart.umd.min.js',
  './vendor/xlsx-js-style.min.js?v=' + APP_VERSION,
  './vendor/qrcode-generator.js',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES)).catch(()=>{})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async ()=>{
    const keys = await caches.keys();
    await Promise.all(
      keys
        .filter((key) => key.startsWith('asistencia-qr-shell-') && key !== CACHE_NAME)
        .map((key) => caches.delete(key))
    );

    await self.clients.claim();

    // Recargar automáticamente toda pestaña abierta de esta plataforma.
    const windows = await self.clients.matchAll({ type:'window', includeUncontrolled:true });
    await Promise.all(windows.map((client)=>{
      try{
        const target = new URL(client.url);
        if(target.origin !== self.location.origin) return Promise.resolve();
        if(target.searchParams.get('appVersion') === APP_VERSION) return Promise.resolve();
        target.searchParams.set('appVersion', APP_VERSION);
        return client.navigate(target.toString()).catch(()=>{});
      }catch(e){
        return Promise.resolve();
      }
    }));
  })());
});

function putInCurrentCache(request, response){
  if(!response || !response.ok) return response;
  const copy = response.clone();
  caches.open(CACHE_NAME).then((cache) => cache.put(request, copy)).catch(()=>{});
  return response;
}

self.addEventListener('fetch', (event) => {
  if(event.request.method !== 'GET') return;
  const url = new URL(event.request.url);

  // Los datos de asistencia siempre van directamente al servidor de Google.
  if(url.hostname.includes('script.google.com')) return;

  const networkFirst =
    event.request.mode === 'navigate' ||
    url.pathname.endsWith('/index.html') ||
    url.pathname.endsWith('/app.js') ||
    url.pathname.endsWith('/manifest.webmanifest') ||
    url.pathname.endsWith('/xlsx-js-style.min.js');

  if(networkFirst){
    event.respondWith(
      fetch(event.request, { cache:'no-store' })
        .then((response) => putInCurrentCache(event.request, response))
        .catch(() => caches.match(event.request).then((cached) => {
          if(cached) return cached;
          if(event.request.mode === 'navigate') return caches.match('./index.html');
          throw new Error('Recurso no disponible sin conexión');
        }))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      if(cached) return cached;
      return fetch(event.request)
        .then((response) => putInCurrentCache(event.request, response));
    })
  );
});
