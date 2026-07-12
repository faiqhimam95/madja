// Vercel Serverless Function (Node runtime) — the only piece that actually holds
// the Firebase service account credential, so it must run server-side, never in
// the client bundle. FIREBASE_SERVICE_ACCOUNT is the full service-account JSON
// (as a string) set via `vercel env add FIREBASE_SERVICE_ACCOUNT`. Called by
// sendPushBestEffort() in src/App.jsx, once per target device token — a failure
// here (invalid/expired token, missing env var, etc.) is swallowed by the caller,
// so this only needs to report the outcome, never coordinate retries itself.
import { initializeApp, cert, getApps } from "firebase-admin/app";
import { getMessaging } from "firebase-admin/messaging";

function getApp() {
  if (getApps().length) return getApps()[0];
  const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
  return initializeApp({ credential: cert(serviceAccount) });
}

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });
  const { token, title, body, data } = req.body || {};
  if (!token || !title) return res.status(400).json({ error: "Missing token or title" });
  try {
    getApp();
    await getMessaging().send({
      token,
      notification: { title, body: body || "" },
      data: Object.fromEntries(Object.entries(data || {}).map(([k, v]) => [k, String(v)])),
      webpush: { fcmOptions: { link: "/" } },
    });
    return res.status(200).json({ ok: true });
  } catch (err) {
    return res.status(200).json({ ok: false, error: err.message });
  }
}
