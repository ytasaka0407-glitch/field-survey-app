// FieldSurveyApp/assets/js/modules/excel/import.js
import {
  ensureSingle,
  ensureMulti,
  getOrInitStationData,
  selectedCategories,
  addSharedStation,
  sharedStations,
} from '../state.js';
import { toInputDateString, stationIdFromName } from '../utils.js';
import { getSchemaFor } from '../ui/schemas.js';

// 画面の基本欄として取り込むラベルとキー（並びは取込時スキャン順に依存）
const CORE_EXTRA_FIELDS = [
  { key: 'installType',     label: '新設/既設流用' },
  { key: 'method',          label: '設置方法' },
  { key: 'diagramStatus',   label: '系統図との整合性' },
  { key: 'diagramNgReason', label: 'NG理由' },
  { key: 'details',         label: 'その他調査内容' },
];

function cellToPlainText(val) {
  if (val == null) return '';
  if (typeof val === 'string') return val;
  if (typeof val === 'number') return String(val);
  if (val && typeof val === 'object' && 'result' in val) return cellToPlainText(val.result);
  if (val && typeof val === 'object' && Array.isArray(val.richText)) return val.richText.map(rt => rt.text || '').join('');
  if (val && typeof val === 'object' && ('text' in val || 'hyperlink' in val)) return String(val.text || val.hyperlink || '');
  try { return String(val); } catch { return '';
  }
}

// 画像Buffer → DataURL変換（埋め込み画像フォールバック用）
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

// 列名の文字（例: 'B'）→ 0-based列番号（B=1）
function colLetterToIndex(L) {
  if (typeof L === 'number') return L;
  if (!L || typeof L !== 'string') return 0;
  const up = L.trim().toUpperCase();
  return up.charCodeAt(0) - 65;
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

  // PHOTO_DATA シートから dataUrl を復元（新フォーマット用）
  // PHOTO_DATA: ['recordId','partIndex','dataPart']
  const photoDataMap = new Map(); // recordId -> reconstructed dataUrl
  const wsPhotoData = wb.getWorksheet('PHOTO_DATA');
  if (wsPhotoData) {
    const lastRow = wsPhotoData.lastRow?.number || 0;
    const partsMap = new Map(); // recordId -> [{idx, part}]
    for (let r = 2; r <= lastRow; r++) {
      const recId  = String(wsPhotoData.getCell(`A${r}`).value ?? '').trim();
      const idxRaw = wsPhotoData.getCell(`B${r}`).value;
      const part   = String(wsPhotoData.getCell(`C${r}`).value ?? '');
      if (!recId) continue;
      const idx = Number(idxRaw ?? 0);
      if (!partsMap.has(recId)) partsMap.set(recId, []);
      partsMap.get(recId).push({ idx, part });
    }
    // 再結合
    partsMap.forEach((arr, recId) => {
      arr.sort((a, b) => a.idx - b.idx);
      const dataUrl = arr.map(x => x.part).join('');
      if (dataUrl && dataUrl.startsWith('data:image/')) {
        photoDataMap.set(recId, dataUrl);
      }
    });
  }

  // PHOTOS シートを最優先で取り込み（新/旧フォーマット両対応）
  // 新: ['recordId','type','sheetName','category','station','fileName','caption','imgCol','imgRowStart']
  // 旧: ['type','category','station','fileName','caption','dataUrl']
  const photosSheet = wb.getWorksheet('PHOTOS');
  let photosRestoredViaSheet = false;
  if (photosSheet) {
    // ヘッダ判定
    const hA1 = String(photosSheet.getCell('A1').value ?? '').trim().toLowerCase();
    const hB1 = String(photosSheet.getCell('B1').value ?? '').trim().toLowerCase();
    const hF1 = String(photosSheet.getCell('F1').value ?? '').trim().toLowerCase();
    const isNewFormat = (hA1 === 'recordid' && hB1 === 'type');
    const isOldFormat = (hA1 === 'type' && hF1 === 'dataurl');

    const lastRow = photosSheet.lastRow?.number || 0;

    for (let r = 2; r <= lastRow; r++) {
      let type = '';
      let sheetName = '';
      let category = '';
      let station = '';
      let fileName = '';
      let caption  = '';
      let dataUrl  = '';

      if (isNewFormat) {
        const recordId = String(photosSheet.getCell(`A${r}`).value ?? '').trim();
        type      = String(photosSheet.getCell(`B${r}`).value ?? '').trim();
        sheetName = String(photosSheet.getCell(`C${r}`).value ?? '').trim();
        category  = String(photosSheet.getCell(`D${r}`).value ?? '').trim();
        station   = String(photosSheet.getCell(`E${r}`).value ?? '').trim();
        fileName  = String(photosSheet.getCell(`F${r}`).value ?? '').trim();
        caption   = String(photosSheet.getCell(`G${r}`).value ?? '').trim();
        // 画像データURLは PHOTO_DATA から復元
        dataUrl   = recordId ? (photoDataMap.get(recordId) || '') : '';
      } else if (isOldFormat) {
        type      = String(photosSheet.getCell(`A${r}`).value ?? '').trim();
        category  = String(photosSheet.getCell(`B${r}`).value ?? '').trim();
        station   = String(photosSheet.getCell(`C${r}`).value ?? '').trim();
        fileName  = String(photosSheet.getCell(`D${r}`).value ?? '').trim();
        caption   = String(photosSheet.getCell(`E${r}`).value ?? '').trim();
        dataUrl   = String(photosSheet.getCell(`F${r}`).value ?? '').trim();
      } else {
        // 不明フォーマットはスキップ
        continue;
      }

      // モデル確保
      let model;
      if (type === 'multi') {
        ensureMulti(category, projectDateEl.value || '');
        const stId = stationIdFromName(station || '基地局');
        if (!sharedStations.some(s => s.id === stId)) {
          addSharedStation({ id: stId, name: station || '基地局' });
        }
        model = getOrInitStationData(category, stId, projectDateEl.value || '');
      } else {
        model = ensureSingle(category, projectDateEl.value || '');
      }
      if (!Array.isArray(model.photos)) model.photos = [];

      // 写真追加（caption はキャプション、name はファイル名）
      model.photos.push({ dataUrl, name: fileName, caption });
      selectedCategories.add(category);
      photosRestoredViaSheet = true;
    }
  }

  // カテゴリシートの読み取り（基本欄・追加項目）
  for (const ws of wb.worksheets) {
    if (!ws) continue;
    const name = ws.name;
    if (name === '表紙' || name === '目次' || name === 'PHOTOS' || name === 'PHOTO_DATA') continue;

    // シート種別とカテゴリ名・基地局名の判定
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
    if (!catName) continue;

    // 基本欄（調査日・設置場所）
    const inDate = toInputDateString(ws.getCell('B3').value);
    const locVal = ws.getCell('B4').value;

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
    if (!Array.isArray(model.photos)) model.photos = [];

    // 基本欄（ラベルで取り込み）: 行5〜「追加項目」まで走査
    let detailsSet = false;
    for (let r = 5; r < 220; r++) {
      const aText = String(ws.getCell(`A${r}`).value ?? '').trim();
      if (!aText) continue;
      if (aText === '追加項目') break;
      const v = ws.getCell(`B${r}`).value;
      const valStr = (v ?? '').toString().trim();

      for (const f of CORE_EXTRA_FIELDS) {
        if (aText === (f.label || f.key)) {
          if (f.key === 'installType') {
            model.installType = /既設/.test(valStr) ? 'reuse' : 'new';
          } else if (f.key === 'diagramStatus') {
            model.diagramStatus = /NG/i.test(valStr) ? 'ng' : 'ok';
          } else if (f.key === 'diagramNgReason') {
            model.diagramNgReason = valStr;
          } else if (f.key === 'method') {
            model.method = valStr;
          } else if (f.key === 'details') {
            model.details = valStr;
            detailsSet = true;
          }
          break;
        }
      }
    }
    // 旧レイアウト互換（B5に調査内容がある場合の取り込み）
    if (!detailsSet) {
      const detVal = ws.getCell('B5').value;
      if (detVal != null) model.details = (detVal ?? '').toString();
    }

    // 追加項目の取り込み（スキーマ準拠）
    const schema = getSchemaFor(catName, isMulti ? 'multi' : 'single');
    let extraHeaderRow = null;
    for (let r = 5; r < 220; r++) {
      const a = String(ws.getCell(`A${r}`).value ?? '').trim();
      if (a === '追加項目') { extraHeaderRow = r; break; }
    }
    if (extraHeaderRow != null) {
      let row = extraHeaderRow + 1;
      let emptyCount = 0;
      const FORBIDDEN_KEYS = new Set(['date','location','installType','method','diagramStatus','diagramNgReason','details','photos']);
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

    // PHOTOS/PHOTO_DATAで復元済みなら、埋め込み画像のフォールバックは不要
    if (photosRestoredViaSheet) continue;

    // フォールバック（PHOTOSがない古いファイルのみ）：埋め込み画像から推定
    if (typeof ws.getImages === 'function' && typeof wb.getImage === 'function') {
      const DESC_COL = 'G';
      const ROWS_PER_CAPTION = 11;
      const MIN_CAPTION_ROW  = 8;

      let images = ws.getImages();
      images = Array.isArray(images) ? images.slice() : [];
      images.sort((a, b) => ((Math.floor(a.tl?.row ?? 0)) - (Math.floor(b.tl?.row ?? 0))));

      const lastRowNum = ws.lastRow?.number || 1000;

      for (let i = 0; i < images.length; i++) {
        const meta = images[i];
        const img = wb.getImage(meta.imageId);
        if (!img || !img.buffer) continue;

        const dataUrl = bufferToDataUrl(img.buffer, img.extension);

        const tlRow0 = Math.floor(meta.tl?.row ?? 0);
        let captionStartRow = Math.max(tlRow0 + 1, MIN_CAPTION_ROW);

        const OFFSETS = [-2, -1, 0, 1, 2, 3, 4];
        for (const off of OFFSETS) {
          const r = tlRow0 + 1 + off;
          if (r < MIN_CAPTION_ROW || r > lastRowNum) continue;
          const cell = ws.getCell(`${DESC_COL}${r}`);
          if (cell && cell.isMerged) {
            captionStartRow = r;
            break;
          }
        }

        let caption = '';
        if (captionStartRow >= 1 && captionStartRow <= lastRowNum) {
          let val = ws.getCell(`${DESC_COL}${captionStartRow}`).value;
          let txt = cellToPlainText(val).trim();

          if (!txt) {
            for (let r = captionStartRow + 1; r < captionStartRow + ROWS_PER_CAPTION && r <= lastRowNum; r++) {
              const v = ws.getCell(`${DESC_COL}${r}`).value;
              const t = cellToPlainText(v).trim();
              if (t) { txt = t; break; }
            }
          }
          caption = txt.replace(/\r\n?/g, '\n');
        }

        model.photos.push({ dataUrl, name: '', caption });
      }
    }
  }
}
