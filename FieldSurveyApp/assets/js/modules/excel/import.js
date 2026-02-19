// FieldSurveyApp/assets/js/modules/excel/import.js
import { ensureSingle, ensureMulti, getOrInitStationData, selectedCategories, addSharedStation, sharedStations } from '../state.js';
import { toInputDateString, stationIdFromName } from '../utils.js';
import { getSchemaFor } from '../ui/schemas.js';

const CORE_EXTRA_FIELDS = [
  { key: 'method',           label: '設置方法',            singleOnly: true },
  { key: 'isReuse',          label: '新設/既設流用' },
  { key: 'diagramOK',        label: '系統図との整合性' },
  { key: 'diagramNgReason',  label: 'NG理由' },
];

function cellToPlainText(val) {
  if (val == null) return '';
  if (typeof val === 'string') return val;
  if (typeof val === 'number') return String(val);
  if (val && typeof val === 'object' && 'result' in val) return cellToPlainText(val.result);
  if (val && typeof val === 'object' && Array.isArray(val.richText)) return val.richText.map(rt => rt.text || '').join('');
  if (val && typeof val === 'object' && ('text' in val || 'hyperlink' in val)) return String(val.text || val.hyperlink || '');
  try { return String(val); } catch { return ''; }
}

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

  const DESC_COL_START = 'G';

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

    // 基本欄の追加項目（Single/Multi共通、method は singleOnly）
    for (let r = 7; r < 200; r++) {
      const aText = String(ws.getCell(`A${r}`).value ?? '').trim();
      if (!aText) continue;
      if (aText === '追加項目') break;
      const v = ws.getCell(`B${r}`).value;
      const valStr = (v ?? '').toString().trim();

      // ラベル一致で取り込み
      for (const f of CORE_EXTRA_FIELDS) {
        if (f.singleOnly && isMulti) continue;
        if (aText === (f.label || f.key)) {
          if (f.key === 'isReuse') {
            model.isReuse = /既設/.test(valStr);
          } else if (f.key === 'diagramOK') {
            model.diagramOK = !/NG/i.test(valStr);
          } else if (f.key === 'diagramNgReason') {
            model.diagramNgReason = valStr;
          } else {
            model[f.key] = valStr;
          }
          break;
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
      const FORBIDDEN_KEYS = new Set(['date','location','details','photos','method','isReuse','diagramOK','diagramNgReason']);
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
          model[targetKey] = valStr;
        }
        row++;
      }
    }

    // 画像の取り込み（略：既存ロジック）
    if (typeof ws.getImages === 'function' && typeof wb.getImage === 'function') {
      const ROWS_PER_CAPTION = 11;
      const BLOCK_ROWS       = 11;
      const GAP_ROWS         = 2;

      let images = ws.getImages();
      images = Array.isArray(images) ? images.slice() : [];
      images.sort((a, b) => ((a.tl?.row ?? 0) - (b.tl?.row ?? 0)));

      const lastRowNum = ws.lastRow?.number || 1000;
      let firstCaptionStart = null;
      for (let r = 8; r <= lastRowNum; r++) {
        const txt = cellToPlainText(ws.getCell(`${DESC_COL_START}${r}`).value).trim();
        if (txt) { firstCaptionStart = r; break; }
      }

      for (let i = 0; i < images.length; i++) {
        const meta = images[i];
        const img = wb.getImage(meta.imageId);
        if (!img || !img.buffer) continue;
        const dataUrl = bufferToDataUrl(img.buffer, img.extension);

        let caption = '';
        if (firstCaptionStart != null) {
          const r0 = firstCaptionStart + i * (BLOCK_ROWS + GAP_ROWS);
          let val = ws.getCell(`${DESC_COL_START}${r0}`).value;
          let txt = cellToPlainText(val).trim();
          if (!txt) {
            for (let r = r0 + 1; r < r0 + ROWS_PER_CAPTION && r <= lastRowNum; r++) {
              const v = ws.getCell(`${DESC_COL_START}${r}`).value;
              const t = cellToPlainText(v).trim();
              if (t) { txt = t; break; }
            }
          }
          caption = txt.replace(/\r\n?/g, '\n');
        }
        model.photos.push({ dataUrl, name: '', caption });
      }
    }
  });

  const photosSheet = wb.getWorksheet('PHOTOS');
  if (photosSheet) {
    const lastRow = photosSheet.lastRow?.number || 0;
    for (let r = 2; r <= lastRow; r++) {
      const type    = String(photosSheet.getCell(`A${r}`).value ?? '').trim();
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
