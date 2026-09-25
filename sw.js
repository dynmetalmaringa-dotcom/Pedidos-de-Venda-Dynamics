/* Pedidos de Venda Dynamics — necessário para o navegador oferecer "Instalar app". Não guarda dados. */
self.addEventListener('install', function (e) { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function () {});
