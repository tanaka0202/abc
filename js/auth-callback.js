// サインイン後に戻ってくるページ（auth.html）。結果を元の画面へ渡して閉じる。
const data = {};
for (const [k, v] of new URLSearchParams(location.search)) data[k] = v;
for (const [k, v] of new URLSearchParams(location.hash.replace(/^#/, ''))) data[k] = v;
const msg = { type: 'mtgrec-auth', data };

// URLに残ったコード・トークンを消す
history.replaceState(null, '', location.pathname);

let framed = false;
try { framed = window.parent !== window; } catch { framed = true; }

if (framed) {
  // 自動のつなぎ直し（アプリの画面の中の、見えない枠）：アプリ本体にだけ結果を渡す。保存も移動もしない
  try { window.parent.postMessage(msg, location.origin); } catch { /* 無視 */ }
} else {
  try { const bc = new BroadcastChannel('mtgrec-auth'); bc.postMessage(msg); bc.close(); } catch { /* 非対応 */ }
  try { localStorage.setItem('mtgrec-auth-result', JSON.stringify(msg)); } catch { /* 無視 */ }
  try { if (window.opener) window.opener.postMessage(msg, location.origin); } catch { /* 無視 */ }
  if (window.opener) {
    setTimeout(() => window.close(), 300);
  } else {
    location.replace(new URL('./#auth-return', location.href).href);
  }
}
