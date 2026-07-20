import React from "react";
import ReactDOM from "react-dom/client";
import { Capacitor } from "@capacitor/core";
import App from "./App.jsx";
import "./theme.css";

// Offline app-shell cache — native Android app only (see public/offline-sw.js for why:
// it's the only context that needs one, and the only one guaranteed not to collide with
// firebase-messaging-sw.js's own root-scope registration on the plain web path).
if (Capacitor.isNativePlatform() && "serviceWorker" in navigator) {
  navigator.serviceWorker.register("/offline-sw.js").catch(() => {});
}

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
