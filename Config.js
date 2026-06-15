/* ─────────────────────────────────────────────────────────────
   NPPS WCR Review — configuration  (wired to your existing setup)
   ───────────────────────────────────────────────────────────── */
window.WCRR_CONFIG = {
  // Same Google project as the current WCR tool.
  GOOGLE_CLIENT_ID: "190605798710-5cashes032781tifqemjuvsm6rvon10c.apps.googleusercontent.com",

  // Your existing Cloudflare Worker (polished-lake) — now also serves
  // /coverage-check and /sentence-grammar.
  WORKER_URL: "https://polished-lake-4911.radheya-supnekar.workers.dev",

  // Live mode. Set to true any time to demo with the built-in sample.
  DEMO_MODE: false,
};

// safety: if either value is blank, fall back to the clickable demo
if (!window.WCRR_CONFIG.GOOGLE_CLIENT_ID || !window.WCRR_CONFIG.WORKER_URL) {
  window.WCRR_CONFIG.DEMO_MODE = true;
}
