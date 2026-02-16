// FieldSurveyApp/assets/js/modules/excel/import.js
import { ensureSingle, ensureMulti, getOrInitStationData, selectedCategories, addSharedStation, sharedStations } from '../state.js';
import { toInputDateString, stationIdFromName } from '../utils.js';
import { getSchemaFor } from '../ui/schemas.js';

// 単一カテゴリで“基本欄”として上部に出力した項目（export.js と対応させる）
const CORE_SINGLE_EXTRA_FIELDS = [
  { key: 'method', label: '設置方法' },
];

function cellToPlainText(val) {
  if (val == null) return '';
  if (typeof val === 'string') return val;
  if (typeof val === 'number') return String(val);
  // 計算式の結果
  if (val && typeof val === 'object' && 'result' in val) {
    return cellToPlainText(val.result);
  }
  // RichText（Excelで太字等を含む場合）
  if (val && typeof val === 'object' && Array.isArray(val.richText)) {
    return val.richText.map(rt => rt.text || '').join('');
  }
  // hyperlink等
  if (val && typeof val === 'object' && ('text' in val || 'hyperlink' in val)) {
    return String(val.text || val.hyperlink || '');
  }
  try { return String(val); } catch { return ''; }
}

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

    // 画像の取り込み（Excel埋め込み画像を読む：画像順にG列の説明ブロックを対応付け）
    if (typeof ws.getImages === 'function' && typeof wb.getImage === 'function') {
      const DESC_COL = 'G';
    
      // ここは「説明ブロックの高さ（行数）」と「ブロック間の余白」に合わせて
      // export 側の設定と同じ値にしてください。
      // 例）説明は11行ぶん、ブロック間に2行の余白 → ROWS_PER_CAPTION=11, BLOCK_ROWS=11, GAP_ROWS=2
      const ROWS_PER_CAPTION = 11;  // 説明として確保している行数
      const BLOCK_ROWS       = 11;  // 1画像ブロックの縦方向行数（説明部の行数に合わせる）
      const GAP_ROWS         = 2;   // ブロック間の余白行数（export側の startRow 増分と一致）
    
      // 1) シート上の画像メタを取得し、縦位置でソート（上→下）
      let images = ws.getImages();
      images = Array.isArray(images) ? images.slice() : [];
      images.sort((a, b) => ((a.tl?.row ?? 0) - (b.tl?.row ?? 0)));
    
      // 2) 最初の説明ブロックの開始行をG列で検出（基本欄を避けるため最低行=8）
      const lastRowNum = ws.lastRow?.number || 1000;
      let firstCaptionStart = null;
      for (let r = 8; r <= lastRowNum; r++) {
        const txt = cellToPlainText(ws.getCell(`${DESC_COL}${r}`).value).trim();
        if (txt) { firstCaptionStart = r; break; }
      }
    
      // 3) 画像と説明を画像順に対応付けて photos へ追加
      for (let i = 0; i < images.length; i++) {
        const meta = images[i];
        const img = wb.getImage(meta.imageId);
        if (!img || !img.buffer) continue;
    
        // DataURLへ変換
        const dataUrl = bufferToDataUrl(img.buffer, img.extension);
    
        // 対応する説明ブロックの開始行を算出
        let caption = '';
        if (firstCaptionStart != null) {
          const r0 = firstCaptionStart + i * (BLOCK_ROWS + GAP_ROWS);
    
          // まずはブロックの先頭行（結合セルのマスター）を読む
          let val = ws.getCell(`${DESC_COL}${r0}`).value;
          let txt = cellToPlainText(val).trim();
    
          // 先頭行が空なら、ブロック内（r0 ～ r0+ROWS_PER_CAPTION-1）で最初にテキストがある行を拾う
          if (!txt) {
            for (let r = r0 + 1; r < r0 + ROWS_PER_CAPTION && r <= lastRowNum; r++) {
              const v = ws.getCell(`${DESC_COL}${r}`).value;
              const t = cellToPlainText(v).trim();
              if (t) { txt = t; break; }
            }
          }
    
          // セル内改行はそのまま維持（CRLF/CRをLFへ正規化）
          caption = txt.replace(/\r\n?/g, '\n');
        }
    
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




