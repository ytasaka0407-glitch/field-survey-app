// FieldSurveyApp/assets/js/modules/pwa.js
export function registerServiceWorker() {
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      // 修正点: 実ファイルの場所に合わせてパスを変更
      navigator.serviceWorker
        .register('./FieldSurveyApp/service-worker.js')
        .then(reg => console.log('Service Worker registered:', reg.scope))
        .catch(err => console.error('Service Worker registration failed:', err));
    });
  }
}
