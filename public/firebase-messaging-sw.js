// Firebase Cloud Messaging background handler — must live at this exact path
// (site root) per FCM's requirements. Config values here are the public client
// config (apiKey etc.), not secrets — same values embedded in every Firebase web
// app, safe to hardcode in a static file the service worker can load without a
// build step. Handles pushes that arrive while the tab isn't focused; foreground
// pushes are handled by src/firebaseClient.js's onMessage listener instead.
importScripts("https://www.gstatic.com/firebasejs/10.14.1/firebase-app-compat.js");
importScripts("https://www.gstatic.com/firebasejs/10.14.1/firebase-messaging-compat.js");

firebase.initializeApp({
  apiKey: "AIzaSyCVRKlIsG4asyMaGqqEkVNYkYLSx6nn7kc",
  authDomain: "sistem-nilai-mdt.firebaseapp.com",
  projectId: "sistem-nilai-mdt",
  storageBucket: "sistem-nilai-mdt.firebasestorage.app",
  messagingSenderId: "1055407163265",
  appId: "1:1055407163265:web:31de1b7723d60e75350a80",
});

const messaging = firebase.messaging();

messaging.onBackgroundMessage((payload) => {
  const { title, body } = payload.notification || {};
  self.registration.showNotification(title || "Sistem Nilai MDT", {
    body: body || "",
    icon: "/kop.png",
    data: payload.data || {},
  });
});
