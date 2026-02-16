// FieldSurveyApp/assets/js/modules/excel/import.js
import { ensureSingle, ensureMulti, getOrInitStationData, selectedCategories, addSharedStation, sharedStations } from '../state.js';
import { toInputDateString, stationIdFromName } from '../utils.js';
import { getSchemaFor } from '../ui/schemas.js';

// 単一カテゴリで“基本欄”として上部に出力した項目（export.js と対応させる）
const CORE_SINGLE_EXTRA_FIELDS = [
  { key: 'method', label: '設置方法' },
];

// 画像Buffer → DataURL変換ヘルパー
function bufferToDataUrl(buffer, extension) {
  const mime = extension === 'png' ? 'image/png' : 'image/jpeg';
  const u8 = (buffer instanceof Uint8Array) ? buffer : new Uint8Array(buffer);
  let binary = '';
  const chunk = 0x8000;
  for (let i = 0; i < u8.length; i += chunk) {
    binary += String.fromCharCode.apply(null, u8.subarray(i, i + chunk));
  }
  const base64 = btoa(binary);
  return `data:${mime};base64,${base64}`;
}

export async function importFromExcel(file, projectTitleEl, projectDateEl, setProjectDatePrev) {
  const buf = await file.arrayBuffer();
  const wb = new window.ExcelJS.Workbook();
  await wb.xlsx.load(buf);

  // 表紙
  const cover = wb.getWorksheet('表紙');
  if (cover) {
    const titleCell = cover.getCell('H8').value;
    const dateCell  = cover.getCell('H10').value;
    projectTitleEl.value = (titleCell ?? '').toString();
    const inDate = toInputDateString(dateCell);
    if (inDate) {
      projectDateEl.value = inDate;
      setProjectDatePrev(inDate);
    }
  }

  // 画像説明カラム（左右分割レイアウトの右側開始列。export.jsのDESC_COL_STARTと一致させる）
  const DESC_COL_START = 'G';

  // カテゴリシート
  wb.worksheets.forEach((ws) => {
    if (!ws) return;
    const name = ws.name;
    if (name === '表紙' || name === '目次' || name === 'PHOTOS') return;

    const label = String(ws.getCell('A1').value ?? '').trim();
    const sheetName = String(ws.name ?? '').trim();

    let isMulti = false;
    let catName = '';
    let stationName = '';

    const mLabel = label.match(/^(.*?)\s*（\s*(.*?)\s*）$/);
    if (mLabel) {
      isMulti     = true;
      catName     = mLabel[1].trim();
      stationName = mLabel[2].trim();
    } else {
      const mName = sheetName.match(/^(.*?)\s*-\s*(.*)$/);
      if (mName) {
        isMulti     = true;
        catName     = mName[1].trim();
        stationName = mName[2].trim();
      } else {
        catName = label || sheetName || '';
      }
    }
    if (!catName) return;

    // 基本項目
    const inDate = toInputDateString(ws.getCell('B3').value);
    const locVal = ws.getCell('B4').value;
    const detVal = ws.getCell('B5').value;

    let model;
    if (isMulti) {
      ensureMulti(catName, projectDateEl.value || "");
      const stId = stationIdFromName(stationName || '基地局');
      if (!sharedStations.some(s => s.id === stId)) {
        addSharedStation({ id: stId, name: stationName || '基地局' });
      }
      model = getOrInitStationData(catName, stId, projectDateEl.value || "");
      selectedCategories.add(catName);
    } else {
      model = ensureSingle(catName, projectDateEl.value || "");
      selectedCategories.add(catName);
    }

    if (inDate) model.date = inDate;
    model.location = (locVal ?? '').toString();
    model.details  = (detVal ?? '').toString();
    if (!Array.isArray(model.photos)) model.photos = [];

    // 単一カテゴリの“基本欄追加”読み取り
    if (!isMulti && CORE_SINGLE_EXTRA_FIELDS.length) {
      for (const f of CORE_SINGLE_EXTRA_FIELDS) {
        for (let r = 7; r < 200; r++) {
          const aText = String(ws.getCell(`A${r}`).value ?? '').trim();
          if (!aText) continue;
          if (aText === '追加項目') break;
          if (aText === (f.label || f.key)) {
            const v = ws.getCell(`B${r}`).value;
            model[f.key] = (v ?? '').toString();
            break;
          }
        }
      }
    }

    // 追加項目の取り込み（既存ロジック）
    const schema = getSchemaFor(catName, isMulti ? 'multi' : 'single');
    let extraHeaderRow = null;
    for (let r = 7; r < 200; r++) {
      const a = String(ws.getCell(`A${r}`).value ?? '').trim();
      if (a === '追加項目') { extraHeaderRow = r; break; }
    }
    if (extraHeaderRow != null) {
      let row = extraHeaderRow + 1;
      let emptyCount = 0;
      const FORBIDDEN_KEYS = new Set(['date','location','details','photos']);
      while (row < extraHeaderRow + 1 + 200) {
        const labelCell = ws.getCell(`A${row}`);
        const valueCell = ws.getCell(`B${row}`);
        const keyCell   = ws.getCell(`K${row}`);
        const labelText = String(labelCell.value ?? '').trim();
        const valRaw    = valueCell.value;
        const keyText   = String(keyCell.value ?? '').trim();
        if (!labelText && !valRaw && !keyText) {
          emptyCount++;
          if (emptyCount >= 3) break;
          row++;
          continue;
        }
        emptyCount = 0;
        let targetKey = keyText;
        if (!targetKey) {
          const f = schema.find(s => (s.label || s.key) === labelText);
          targetKey = f?.key || '';
        }
        if (targetKey && !FORBIDDEN_KEYS.has(targetKey)) {
          const valStr = (valRaw ?? '').toString();
          if (targetKey === 'method' && model[targetKey]) {
            // 既に基本欄から取得済みならスキップ
          } else {
            model[targetKey] = valStr;
          }
        }
        row++;
      }
    }

    // 画像の取り込み（Excel埋め込み画像を読む）
    // ExcelJSのビルドが getImages / getImage をサポートしている場合のみ動作
    if (typeof ws.getImages === 'function' && typeof wb.getImage === 'function') {
      const images = ws.getImages();
      for (const meta of images) {
        const img = wb.getImage(meta.imageId);
        if (!img || !img.buffer) continue;
        // 画像のtop-left行（exportのtl.rowは0始まりなので+1）
        const startRow = Math.round((meta.tl?.row ?? 0) + 1);
        // 説明は右側カラムの開始セル（DESC_COL_START）にある前提
        const capCell = ws.getCell(`${DESC_COL_START}${startRow}`).value;
        const caption = (capCell ?? '').toString();
        const dataUrl = bufferToDataUrl(img.buffer, img.extension);
        model.photos.push({ dataUrl, name: '', caption });
      }
    }
  });

  // 代替：隠しシート "PHOTOS" から写真復元（ExcelJSが画像読めない場合の保険）
  const photosSheet = wb.getWorksheet('PHOTOS');
  if (photosSheet) {
    // 期待する列: A=type, B=category, C=station, D=fileName, E=caption, F=dataUrl
    const lastRow = photosSheet.lastRow?.number || 0;
    for (let r = 2; r <= lastRow; r++) {
      const type    = String(photosSheet.getCell(`A${r}`).value ?? '').trim(); // 'single' / 'multi'
      const cat     = String(photosSheet.getCell(`B${r}`).value ?? '').trim();
      const station = String(photosSheet.getCell(`C${r}`).value ?? '').trim();
      const name    = String(photosSheet.getCell(`D${r}`).value ?? '').trim();
      const caption = String(photosSheet.getCell(`E${r}`).value ?? '').trim();
      const dataUrl = String(photosSheet.getCell(`F${r}`).value ?? '').trim();
      if (!cat || !dataUrl) continue;

      let model;
      if (type === 'multi') {
        ensureMulti(cat, projectDateEl.value || "");
        const stId = stationIdFromName(station || '基地局');
        if (!sharedStations.some(s => s.id === stId)) {
          addSharedStation({ id: stId, name: station || '基地局' });
        }
        model = getOrInitStationData(cat, stId, projectDateEl.value || "");
      } else {
        model = ensureSingle(cat, projectDateEl.value || "");
      }
      if (!Array.isArray(model.photos)) model.photos = [];
      model.photos.push({ dataUrl, name, caption });
      selectedCategories.add(cat);
    }
  }
}
