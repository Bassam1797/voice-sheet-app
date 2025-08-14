self.addEventListener('install', (e)=>{
  e.waitUntil(caches.open('vs-cache-v2').then(cache=>cache.addAll([
    './','./index.html','./assets/style.css',
    './src/grid.js','./src/speech.js','./src/export.js','./src/storage.js','./src/utils.js',
    'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'
  ])));
});
self.addEventListener('fetch', (e)=>{
  e.respondWith(caches.match(e.request).then(resp=>resp || fetch(e.request)));
});
