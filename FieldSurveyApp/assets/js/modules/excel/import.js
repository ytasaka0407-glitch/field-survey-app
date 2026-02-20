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
  try { return String(val); } catch { return ''; }
}

// フォールバック用（PHOTOSが無い古いファイル向け／画像埋め込みから dataUrl を生成）
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

  // PHOTO_DATA（新方式）: ['recordId','partIndex','dataPart'] を再結合
  const photoDataMap = new Map(); // recordId -> dataUrl
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
    for (const [recId, arr] of partsMap.entries()) {
      arr.sort((a, b) => a.idx - b.idx);
      const dataUrl = arr.map(x => x.part).join('');
      if (dataUrl && dataUrl.startsWith('data:image/')) {
        photoDataMap.set(recId, dataUrl);
      }
    }
  }

  // PHOTOS 読み込み（新/旧/マニフェストのみ対応）
  // 新: ['recordId','type','sheetName','category','station','fileName','caption','imgCol','imgRowStart']
  // 旧: ['type','category','station','fileName','caption','dataUrl']
  // マニフェストのみ（本アプリのエクスポート）: ['type','sheetName','category','station','fileName','caption','imgCol','imgRowStart']
  const photosSheet = wb.getWorksheet('PHOTOS');
  const photosBySheet = new Map();      // sheetName -> [{ type, category, station, name, caption, dataUrl, imgRowStart, recordId? }]
  const photosByKey = new Map();        // `${type}|${category}|${station}` -> [{...}] （旧フォーマット用）

  if (photosSheet) {
    const hA1 = String(photosSheet.getCell('A1').value ?? '').trim().toLowerCase();
    const hB1 = String(photosSheet.getCell('B1').value ?? '').trim().toLowerCase();
    const hC1 = String(photosSheet.getCell('C1').value ?? '').trim().toLowerCase();
    const hD1 = String(photosSheet.getCell('D1').value ?? '').trim().toLowerCase();
    const hE1 = String(photosSheet.getCell('E1').value ?? '').trim().toLowerCase();
    const hF1 = String(photosSheet.getCell('F1').value ?? '').trim().toLowerCase();
    const hG1 = String(photosSheet.getCell('G1').value ?? '').trim().toLowerCase();
    const hH1 = String(photosSheet.getCell('H1').value ?? '').trim().toLowerCase();
    const hI1 = String(photosSheet.getCell('I1').value ?? '').trim().toLowerCase();

    const isNewFormat       = (hA1 === 'recordid' && hB1 === 'type');
    const isOldFormat       = (hA1 === 'type' && hF1 === 'dataurl');
    const isManifestFormat  = (hA1 === 'type' && hB1 === 'sheetname' && hC1 === 'category' && hD1 === 'station' &&
                               hE1 === 'filename' && hF1 === 'caption' && hG1 === 'imgcol' && hH1 === 'imgrowstart');

    const lastRow = photosSheet.lastRow?.number || 0;
    for (let r = 2; r <= lastRow; r++) {
      if (isNewFormat) {
        const recordId   = String(photosSheet.getCell(`A${r}`).value ?? '').trim();
        const type       = String(photosSheet.getCell(`B${r}`).value ?? '').trim();
        const sheetName  = String(photosSheet.getCell(`C${r}`).value ?? '').trim();
        const category   = String(photosSheet.getCell(`D${r}`).value ?? '').trim();
        const station    = String(photosSheet.getCell(`E${r}`).value ?? '').trim();
        const fileName   = String(photosSheet.getCell(`F${r}`).value ?? '').trim();
        const caption    = String(photosSheet.getCell(`G${r}`).value ?? '').trim();
        const imgRowRaw  = photosSheet.getCell(`I${r}`).value;
        const imgRowStart = Number(imgRowRaw || 0);
        const dataUrl    = recordId ? (photoDataMap.get(recordId) || '') : '';
        const rec = { type, category, station, name: fileName, caption, dataUrl, imgRowStart, recordId };
        if (!photosBySheet.has(sheetName)) photosBySheet.set(sheetName, []);
        photosBySheet.get(sheetName).push(rec);
      } else if (isOldFormat) {
        const type      = String(photosSheet.getCell(`A${r}`).value ?? '').trim();
        const category  = String(photosSheet.getCell(`B${r}`).value ?? '').trim();
        const station   = String(photosSheet.getCell(`C${r}`).value ?? '').trim();
        const fileName  = String(photosSheet.getCell(`D${r}`).value ?? '').trim();
        const caption   = String(photosSheet.getCell(`E${r}`).value ?? '').trim();
        const dataUrl   = String(photosSheet.getCell(`F${r}`).value ?? '').trim();
        const key = `${type}|${category}|${station}`;
        const rec = { type, category, station, name: fileName, caption, dataUrl, imgRowStart: 0 };
        if (!photosByKey.has(key)) photosByKey.set(key, []);
        photosByKey.get(key).push(rec);
      } else if (isManifestFormat) {
        const type       = String(photosSheet.getCell(`A${r}`).value ?? '').trim();
        const sheetName  = String(photosSheet.getCell(`B${r}`).value ?? '').trim();
        const category   = String(photosSheet.getCell(`C${r}`).value ?? '').trim();
        const station    = String(photosSheet.getCell(`D${r}`).value ?? '').trim();
        const fileName   = String(photosSheet.getCell(`E${r}`).value ?? '').trim();
        const caption    = String(photosSheet.getCell(`F${r}`).value ?? '').trim();
        const imgRowRaw  = photosSheet.getCell(`H${r}`).value;
        const imgRowStart = Number(imgRowRaw || 0);
        const rec = { type, category, station, name: fileName, caption, dataUrl: '', imgRowStart };
        if (!photosBySheet.has(sheetName)) photosBySheet.set(sheetName, []);
        photosBySheet.get(sheetName).push(rec);
      }
    }
  }

  // カテゴリシート 読み取り
  for (const ws of wb.worksheets) {
    if (!ws) continue;
    const name = ws.name;
    if (name === '表紙' || name === '目次' || name === 'PHOTOS' || name === 'PHOTO_DATA') continue;

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

    // 基本欄
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

    // 基本欄（ラベル取り）
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
            model.details = valStr; detailsSet = true;
          }
          break;
        }
      }
    }
    if (!detailsSet) {
      const detVal = ws.getCell('B5').value;
      if (detVal != null) model.details = (detVal ?? '').toString();
    }

    // 追加項目（スキーマ）
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
          emptyCount++; if (emptyCount >= 3) break; row++; continue;
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

    // 写真（優先順: PHOTOS（新/マニフェスト）×シート画像突合 → PHOTOS旧方式 → 画像フォールバック）
    let sheetPhotosApplied = false;

    // PHOTOS（新/マニフェスト）＋シートの埋め込み画像を「行開始」で1対1突き合わせ
    if (photosBySheet.has(sheetName)) {
      const allRecs = photosBySheet.get(sheetName);
      const recs = allRecs
        .filter(rec =>
          rec.type === (isMulti ? 'multi' : 'single') &&
          rec.category === catName &&
          (isMulti ? (rec.station === (stationName || '')) : true)
        )
        .filter(rec => Number.isFinite(Number(rec.imgRowStart)) && Number(rec.imgRowStart) > 0)
        .map(rec => ({ ...rec, imgRowStart: Number(rec.imgRowStart) }))
        .sort((a, b) => a.imgRowStart - b.imgRowStart);

      if (recs.length) {
        // G列のキャプション（行 → caption）を事前取得
        const captionByRow = new Map();
        for (const rec of recs) {
          const v = ws.getCell(`G${rec.imgRowStart}`).value;
          const t = cellToPlainText(v).trim();
          captionByRow.set(rec.imgRowStart, t ? t.replace(/\r\n?/g, '\n') : '');
        }

        // シートの画像埋め込み一覧（位置取得）
        let images = (typeof ws.getImages === 'function') ? ws.getImages() : [];
        images = Array.isArray(images) ? images.slice() : [];
        // 画像ごとに「推定開始行」を計算
        const imgPool = images.map(meta => {
          const imageObj = (typeof wb.getImage === 'function') ? wb.getImage(meta.imageId) : null;
          if (!imageObj || !imageObj.buffer) return null;
          const tlRow0 = Math.floor(meta.tl?.row ?? 0); // 0-based
          const approxStartRow = tlRow0 + 1;            // 1-based
          return { meta, imageObj, approxStartRow, used: false };
        }).filter(Boolean);

        // 許容ウィンドウ（エクスポート側は 1枚あたり 11 行確保）
        const MATCH_WINDOW = 12; // 行幅の許容（開始行〜+11 付近）
        const candidatesInWindow = (imgRow, recRow) =>
          imgRow >= recRow && imgRow <= (recRow + MATCH_WINDOW - 1);

        for (const rec of recs) {
          // rec.imgRowStart と最も近い未使用画像を選ぶ（同じブロック内優先）
          let bestIdx = -1;
          let bestDist = Number.POSITIVE_INFINITY;

          for (let i = 0; i < imgPool.length; i++) {
            const it = imgPool[i];
            if (it.used) continue;
            if (!candidatesInWindow(it.approxStartRow, rec.imgRowStart)) continue;
            const dist = Math.abs(it.approxStartRow - rec.imgRowStart);
            if (dist < bestDist) { bestDist = dist; bestIdx = i; }
          }

          if (bestIdx >= 0) {
            const it = imgPool[bestIdx];
            it.used = true;

            // キャプション: Gセル優先、空ならPHOTOSのcaption
            let caption = captionByRow.get(rec.imgRowStart) || rec.caption || '';
            if (caption && model.diagramNgReason && caption === model.diagramNgReason && rec.caption) {
              caption = rec.caption;
            }

            const dataUrl = bufferToDataUrl(it.imageObj.buffer, it.imageObj.extension);
            model.photos.push({
              dataUrl,
              name: rec.name || '',
              caption
            });
          }
        }

        sheetPhotosApplied = model.photos.length > 0;
      }
    }

    // PHOTOS（旧方式）
    if (!sheetPhotosApplied && photosByKey.size) {
      const key = `${isMulti ? 'multi' : 'single'}|${catName}|${isMulti ? (stationName || '') : ''}`;
      const recs = photosByKey.get(key);
      if (Array.isArray(recs) && recs.length) {
        for (const rec of recs) {
          model.photos.push({
            dataUrl: rec.dataUrl || '',
            name: rec.name || '',
            caption: rec.caption || ''
          });
        }
        sheetPhotosApplied = true;
      }
    }

    // フォールバック（PHOTOSが無い or 非対応フォーマット）: シート内画像と G列結合キャプションから再構築
    if (!sheetPhotosApplied && typeof ws.getImages === 'function' && typeof wb.getImage === 'function') {
      const DESC_COL = 'G';
      const ROWS_PER_CAPTION = 11;
      const MIN_CAPTION_ROW  = 8;

      let images = ws.getImages();
      images = Array.isArray(images) ? images.slice() : [];
      images.sort((a, b) => ((Math.floor(a.tl?.row ?? 0)) - (Math.floor(b.tl?.row ?? 0))));
      const lastRow = ws.lastRow?.number || 1000;

      for (const meta of images) {
        const imageObj = wb.getImage(meta.imageId);
        if (!imageObj || !imageObj.buffer) continue;
        const dataUrl = bufferToDataUrl(imageObj.buffer, imageObj.extension);

        const tlRow0 = Math.floor(meta.tl?.row ?? 0);
        let start = Math.max(tlRow0 + 1, MIN_CAPTION_ROW);
        const end = Math.min(start + ROWS_PER_CAPTION - 1, lastRow);

        // 近傍の結合セル（G列マスター）を優先し、B列結合（基本欄）は除外
        const OFFSETS = [-2, -1, 0, 1, 2, 3, 4];
        for (const off of OFFSETS) {
          const r = tlRow0 + 1 + off;
          if (r < MIN_CAPTION_ROW || r > lastRow) continue;
          const bCell = ws.getCell(`B${r}`);
          if (bCell && bCell.isMerged) continue;
          const gCell = ws.getCell(`${DESC_COL}${r}`);
          if (gCell && gCell.isMerged) { start = r; break; }
        }

        let caption = '';
        for (let r = start; r <= end; r++) {
          const bCell = ws.getCell(`B${r}`);
          if (bCell && bCell.isMerged) continue;
          const gCell = ws.getCell(`${DESC_COL}${r}`);
          if (gCell && gCell.isMerged) {
            const t = cellToPlainText(gCell.value).trim();
            if (t) { caption = t.replace(/\r\n?/g, '\n'); break; }
          }
        }

        model.photos.push({ dataUrl, name: '', caption });
      }
    }
  }
}
