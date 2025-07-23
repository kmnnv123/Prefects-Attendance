// Service Worker for PWA functionality
const CACHE_NAME = 'bop25-v1.0.0';
const urlsToCache = [
  '/',
  '/index.html',
  '/script.js',
  '/manifest.json',
  '/prefects logo.png',
  '/Picture1.png',
  'https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css',
  'https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css',
  'https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js',
  'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js'
];

// Install Service Worker
self.addEventListener('install', (event) => {
  console.log('📱 PWA: Service Worker installing...');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        console.log('📱 PWA: Opened cache');
        return cache.addAll(urlsToCache);
      })
      .catch((error) => {
        console.log('📱 PWA: Cache failed:', error);
      })
  );
});

// Fetch events
self.addEventListener('fetch', (event) => {
  event.respondWith(
    caches.match(event.request)
      .then((response) => {
        // Return cached version or fetch from network
        if (response) {
          return response;
        }
        return fetch(event.request);
      })
  );
});

// Activate Service Worker
self.addEventListener('activate', (event) => {
  console.log('📱 PWA: Service Worker activating...');
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            console.log('📱 PWA: Deleting old cache:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
});

// Background sync for offline functionality
self.addEventListener('sync', (event) => {
  if (event.tag === 'background-sync') {
    console.log('📱 PWA: Background sync triggered');
    // Add any background sync logic here
  }
});

// Push notifications (if needed later)
self.addEventListener('push', (event) => {
  if (event.data) {
    const data = event.data.json();
    console.log('📱 PWA: Push received:', data);
    
    const options = {
      body: data.body,
      icon: '/prefects logo.png',
      badge: '/prefects logo.png',
      vibrate: [100, 50, 100],
      data: {
        dateOfArrival: Date.now(),
        primaryKey: '2'
      }
    };
    
    event.waitUntil(
      self.registration.showNotification(data.title, options)
    );
  }
});
