export async function ensureExcelLibs() {
  if (window.ExcelJS && window.saveAs) return;

  const loadScript = (src) => new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src; s.async = true;
    s.onload = resolve; s.onerror = reject;
    document.head.appendChild(s);
  });

  await loadScript('./lib/exceljs.min.js');
  await loadScript('./lib/FileSaver.min.js');
}