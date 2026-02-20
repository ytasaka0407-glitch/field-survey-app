// FieldSurveyApp/assets/js/modules/pwa.js
export function registerServiceWorker() {
  if (!('serviceWorker' in navigator)) return;

  window.addEventListener('load', () => {
    navigator.serviceWorker
      .register('./service-worker.js') // ルートのSWを登録
      .then((reg) => {
        console.log('Service Worker registered:', reg.scope);

        // 新しいSWが waiting になったら更新確認 → 即時適用
        function promptUpdate(r) {
          if (r.waiting) {
            const ok = confirm('新しいバージョンがあります。更新しますか？');
            if (ok) r.waiting.postMessage({ type: 'SKIP_WAITING' });
          }
        }

        // 既に waiting がいる場合（初回登録直後にあり得る）
        promptUpdate(reg);

        // 更新検知（updatefound → installed → waiting）
        reg.addEventListener('updatefound', () => {
          const nw = reg.installing;
          if (!nw) return;
          nw.addEventListener('statechange', () => {
            if (nw.state === 'installed') {
              // 旧SWが制御中なら waiting → ユーザーに更新確認
              if (navigator.serviceWorker.controller) {
                promptUpdate(reg);
              }
            }
          });
        });

        // SWが切り替わったら自動リロード（更新反映）
        navigator.serviceWorker.addEventListener('controllerchange', () => {
          // 一度だけリロード
          if (!window.__swReloaded) {
            window.__swReloaded = true;
            location.reload();
          }
        });
      })
      .catch((err) => console.error('Service Worker registration failed:', err));
  });
}
