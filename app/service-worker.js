// Назва кешу — при зміні файлів додатку, міняти версію (v2, v3...)
// щоб телефони підхопили оновлення
const CACHE_NAME = "tempo-v1";

// Всі файли які треба зберегти офлайн
const FILES_TO_CACHE = [
  "./index.html",
  "./screen2.html",
  "./screen3.html",
  "./manifest.json",
  "images/background2.png",
  "icons/CP.png"
];

// ПОДІЯ 1: install — спрацьовує один раз коли service worker вперше завантажується
// Тут ми відкриваємо кеш і зберігаємо всі файли додатку
self.addEventListener("install", function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(FILES_TO_CACHE);
    })
  );
});

// ПОДІЯ 2: fetch — спрацьовує КОЖНОГО РАЗУ коли браузер щось запитує
// (відкриває сторінку, завантажує файл тощо)
// Ми перехоплюємо запит і відповідаємо з кешу замість інтернету
self.addEventListener("fetch", function(event) {
  event.respondWith(
    caches.match(event.request).then(function(response) {
      // якщо файл є в кеші — повертаємо його (офлайн працює!)
      // якщо немає — пробуємо інтернет (fallback)
      return response || fetch(event.request);
    })
  );
});

// ПОДІЯ 3: activate — спрацьовує коли встановлюється нова версія
// Тут видаляємо старий кеш щоб не було конфліктів
self.addEventListener("activate", function(event) {
  event.waitUntil(
    caches.keys().then(function(keyList) {
      return Promise.all(keyList.map(function(key) {
        if (key !== CACHE_NAME) {
          return caches.delete(key);
        }
      }));
    })
  );
});
