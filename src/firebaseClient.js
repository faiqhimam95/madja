// Firebase Cloud Messaging client — Android only for now, per explicit scope
// decision (iOS deferred, needs Mac). Two delivery paths share the same
// deviceTokens slice + api/send-push.js server: plain browser (Chrome Web Push,
// via firebase/messaging) and the Capacitor-wrapped native app (native FCM via
// @capacitor/push-notifications) — a token is a token to the Admin SDK either
// way. This same JS bundle runs in both contexts since capacitor.config.json
// points server.url at the live madja.vercel.app deploy rather than bundling a
// separate copy; Capacitor.isNativePlatform() is what tells them apart at
// runtime. firebaseConfig values are the public client config (not secret) —
// same ones baked into public/firebase-messaging-sw.js and
// android/app/google-services.json, both fetched automatically via `firebase
// apps:sdkconfig` during setup rather than copy-pasted from the Console.
// VAPID_KEY is the one piece Firebase only lets you generate/view from the
// Console UI (Project Settings > Cloud Messaging > Web Push certificates) — no
// CLI shortcut exists for it, and it's only used on the web path (native FCM
// doesn't need it).
import { initializeApp } from "firebase/app";
import { getMessaging, getToken, onMessage, isSupported } from "firebase/messaging";
import { Capacitor } from "@capacitor/core";
import { PushNotifications } from "@capacitor/push-notifications";

const firebaseConfig = {
  apiKey: "AIzaSyCVRKlIsG4asyMaGqqEkVNYkYLSx6nn7kc",
  authDomain: "sistem-nilai-mdt.firebaseapp.com",
  projectId: "sistem-nilai-mdt",
  storageBucket: "sistem-nilai-mdt.firebasestorage.app",
  messagingSenderId: "1055407163265",
  appId: "1:1055407163265:web:31de1b7723d60e75350a80",
};

const VAPID_KEY = "BE8S-bdm2aryWCxoP066nCgv3hUmo04lkMw4_I1oBMrAOnbNrZXBb7W4hf3ahQed4Z89zAPrvyYco3DduasnCA0";

let appPromise = null;
function getFirebaseApp() {
  if (!appPromise) appPromise = Promise.resolve(initializeApp(firebaseConfig));
  return appPromise;
}

// Requests notification permission (if not already decided) and registers this
// device's FCM token via `onToken`. Fully self-contained no-op on unsupported
// browsers/contexts (non-HTTPS, no Notification API, VAPID key not yet filled
// in, user declines) — every failure mode here must be silent, this is an
// enhancement layered on top of the in-app notifikasi bell, never a hard
// requirement for the rest of the app to function.
export async function requestPushPermission(onToken) {
  if (Capacitor.isNativePlatform()) return requestNativePushPermission(onToken);
  return requestWebPushPermission(onToken);
}

async function requestNativePushPermission(onToken) {
  try {
    const perm = await PushNotifications.requestPermissions();
    if (perm.receive !== "granted") return false;
    await PushNotifications.addListener("registration", (token) => onToken(token.value));
    await PushNotifications.register();
    return true;
  } catch {
    return false;
  }
}

async function requestWebPushPermission(onToken) {
  try {
    if (VAPID_KEY.startsWith("REPLACE_WITH")) return false;
    if (!("Notification" in window) || !(await isSupported())) return false;
    const permission = await Notification.requestPermission();
    if (permission !== "granted") return false;
    const app = await getFirebaseApp();
    const registration = await navigator.serviceWorker.register("/firebase-messaging-sw.js");
    const messaging = getMessaging(app);
    const token = await getToken(messaging, { vapidKey: VAPID_KEY, serviceWorkerRegistration: registration });
    if (token) onToken(token);
    onMessage(messaging, (payload) => {
      const { title, body } = payload.notification || {};
      if (title && Notification.permission === "granted") new Notification(title, { body, icon: "/kop.png" });
    });
    return true;
  } catch {
    return false;
  }
}
