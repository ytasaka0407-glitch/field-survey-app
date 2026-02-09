// assets/js/main.js (module)
import { registerServiceWorker } from './modules/pwa.js';
import { bootstrapApp } from './app.js';

registerServiceWorker();
window.addEventListener('DOMContentLoaded', () => {
  bootstrapApp();
});