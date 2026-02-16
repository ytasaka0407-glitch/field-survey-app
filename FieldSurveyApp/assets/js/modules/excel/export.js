// FieldSurveyApp/assets/js/modules/excel/export.js
import { dataMap, selectedCategories, ensureSingle, ensureMulti, getOrInitStationData } from '../state.js';
import { sanitizeSheetName, makeUniqueSheetName, fromInputDate, getImageDim } from '../utils.js';
import { getSchemaFor } from '../ui/schemas.js';

export async function exportToExcel(projectTitle, projectDate, sharedStations) {
  const ExcelJSRef = window.ExcelJS;
  const wb = new ExcelJSRef.Workbook();
  wb.creator = '現地調査レポートツール';
  wb.created = new Date();
  wb.properties.title = '現地調査報告書';

  const titleStyle        = { font: { size: 20, bold: true, name: 'Meiryo UI' }, alignment: { horizontal: 'center', vertical: 'middle' } };
  const sectionTitleStyle = { font: { size: 14, bold: true, name: 'Meiryo UI' } };
  const labelStyle        = { font: { bold: true, name: 'Meiryo UI' } };
  const linkStyle         = { font: { color: { argb: 'FF1F4E79' }, underline: true, name: 'Meiryo UI' } };
  const borderThin        = { style: 'thin', color: { argb: 'FF999999' } };

  // 単一カテゴリの“基本欄”として上部に出力したい追加項目（ここに追記していく）
  const CORE_SINGLE_EXTRA_FIELDS = [
    { key: 'method', label: '設置方法' }, // ← 追加したい項目
  ];

  const projectDateStr = projectDate || '';
  const exportDate     = new Date();
  const coverDate      = fromInputDate(projectDateStr) || exportDate;

  // 表紙
  const wsCover = wb.addWorksheet('表紙', {
    pageSetup: { paperSize: 9, orientation: 'portrait', margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 } },
    headerFooter: { oddFooter: '&Rページ &P / &N' }
  });
  for (let i = 1; i <= 12; i++) wsCover.getColumn(i).width = 16;
  wsCover.mergeCells('A3:H6');
  wsCover.getCell('A3').value = '現地調査報告書';
  wsCover.getCell('A3').style = titleStyle;
  wsCover.getCell('G8').value = '案件名'; wsCover.getCell('G8').style = labelStyle;
  wsCover.getCell('H8').value = projectTitle || '-';
  wsCover.getCell('H8').font  = { name: 'Meiryo UI' };
  wsCover.getCell('G10').value = '日付'; wsCover.getCell('G10').style = labelStyle;
  wsCover.getCell('H10').value = coverDate; wsCover.getCell('H10').numFmt = 'yyyy/mm/dd';
  wsCover.getCell('H10').font = { name: 'Meiryo UI' };
  wsCover.getCell('H10').alignment = { horizontal: 'left' };

  // 目次
  const wsToc = wb.addWorksheet('目次', { pageSetup: { paperSize: 9, orientation: 'portrait' }, headerFooter: { oddFooter: '&Rページ &P / &N' } });
  wsToc.getColumn(1).width = 50;
  wsToc.getColumn(2).width = 18;
  wsToc.getCell('A1').value = '目次';
  wsToc.getCell('A1').style = sectionTitleStyle;

  const cats = [...selectedCategories];
  const selectedSingles = cats.filter(c => (dataMap[c]?.mode || 'single') === 'single');
  const selectedMultis  = cats.filter(c => (dataMap[c]?.mode) === 'multi');

  const entries = [];
  for (const cat of selectedSingles) {
    const v = ensureSingle(cat, projectDateStr);
    entries.push({ type: 'single', cat, stationId: null, stationName: null, displayLabel: cat, model: v });
  }
  for (const cat of selectedMultis) {
    const mv = ensureMulti(cat, projectDateStr);
    const stationList = sharedStations.length ? sharedStations.slice() : Object.keys(mv.stationData || {}).map(id => ({ id, name: id }));
    for (const st of stationList) {
      const stData = getOrInitStationData(cat, st.id, projectDateStr);
      entries.push({ type: 'multi', cat, stationId: st.id, stationName: st.name, displayLabel: `${cat}（${st.name}）`, model: stData });
    }
  }

  const usedNames = new Set();
  for (const e of entries) {
    const base = e.type === 'single' ? sanitizeSheetName(e.cat) : sanitizeSheetName(`${e.cat} - ${e.stationName}`);
    e.sheetName = makeUniqueSheetName(base, usedNames);
    usedNames.add(e.sheetName);
  }

  const colLetterToIndex = (L) => L.charCodeAt(0) - 65;

  async function addOneEntrySheet(entry) {
    const ws = wb.addWorksheet(entry.sheetName, {
      pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 } },
      headerFooter: { oddFooter: '&Rページ &P / &N' }
    });
    ws.views = [{ showGridLines: false }];

    // 列幅設定 + メタキー用のK列は非表示にする
    const colWidths = [10, 16, 16, 18, 16, 16, 18, 16, 16, 18, 10, 10]; // A〜L
    colWidths.forEach((w, i) => ws.getColumn(i+1).width = w);
    ws.getColumn(11).hidden = true; // K列（フィールドキー格納）

    ws.getCell('A1').value = entry.displayLabel;
    ws.getCell('A1').style = sectionTitleStyle;

    // 基本項目（既存3項目）
    ws.getCell('A3').value = '調査日'; ws.getCell('A3').style = labelStyle;
    const d = fromInputDate(entry.model.date || projectDateStr);
    if (d) { ws.getCell('B3').value = d; ws.getCell('B3').numFmt = 'yyyy/mm/dd'; }
    else { ws.getCell('B3').value = '-'; }
    ws.getCell('B3').font = { name: 'Meiryo UI' };
    ws.getCell('B3').alignment = { horizontal: 'left' };

    ws.getCell('A4').value = '設置場所'; ws.getCell('A4').style = labelStyle;
    ws.mergeCells('B4:J4'); ws.getCell('B4').value = entry.model.location || '-';
    ws.getCell('B4').font  = { name: 'Meiryo UI' };

    ws.getCell('A5').value = '調査内容'; ws.getCell('A5').style = labelStyle;
    ws.mergeCells('B5:J6');
    ws.getCell('B5').value     = (entry.model.details || '').replace(/\r?\n/g, '\n');
    ws.getCell('B5').alignment = { wrapText: true, vertical: 'top' };
    ws.getCell('B5').font      = { name: 'Meiryo UI' };

    ws.getCell('A3').border = { bottom: borderThin };
    ws.getCell('A4').border = { bottom: borderThin };
    ws.getCell('A5').border = { bottom: borderThin };
    ws.getCell('B3').border = { bottom: borderThin };
    ws.getCell('B4').border = { bottom: borderThin };
    ws.getCell('B5').border = { bottom: borderThin };

    // ここから“基本欄として出す追加項目（単一カテゴリ用）”
    let nextRow = 7; // A7/B7 から追記
    if (entry.type === 'single' && CORE_SINGLE_EXTRA_FIELDS.length) {
      for (const f of CORE_SINGLE_EXTRA_FIELDS) {
        ws.getCell(`A${nextRow}`).value = f.label;
        ws.getCell(`A${nextRow}`).style = labelStyle;
        ws.mergeCells(`B${nextRow}:J${nextRow}`);
        ws.getCell(`B${nextRow}`).value = (entry.model[f.key] ?? '').toString() || '-';
        ws.getCell(`B${nextRow}`).font = { name: 'Meiryo UI' };
        nextRow++;
      }
    }

    // “追加項目”セクション（スキーマ定義に基づく。コア扱いのキーは除外）
    const schema = getSchemaFor(entry.cat, entry.type === 'multi' ? 'multi' : 'single');
    const CORE_KEYS = new Set(['date','location','details','photos','method']); // ← method をコア扱いで除外
    const extraFields = schema.filter(f => !CORE_KEYS.has(f.key));

    if (extraFields.length) {
      ws.getCell(`A${nextRow}`).value = '追加項目';
      ws.getCell(`A${nextRow}`).style = labelStyle;
      let extraStartRow = nextRow + 1;

      for (let i = 0; i < extraFields.length; i++) {
        const f = extraFields[i];
        const row = extraStartRow + i;
        ws.getCell(`A${row}`).value = f.label || f.key;
        ws.getCell(`A${row}`).font  = { name: 'Meiryo UI' };
        ws.mergeCells(`B${row}:J${row}`);
        const val = entry.model[f.key];
        ws.getCell(`B${row}`).value = (val ?? '').toString();
        ws.getCell(`B${row}`).alignment = { wrapText: true, vertical: 'top' };
        ws.getCell(`B${row}`).font = { name: 'Meiryo UI' };
        ws.getCell(`K${row}`).value = f.key; // インポート用キー
      }
      nextRow = extraStartRow + extraFields.length;
    }

    // 写真セクション（左右分割：左=画像、右=説明。1ページ2枚相当）
    let startRow = nextRow + 1;
    const photos = Array.isArray(entry.model.photos) ? entry.model.photos : [];
    
    // 左側の画像エリアの列（B〜F）、右側の説明エリアの列（G〜J）
    const IMAGE_COLS = ['B','C','D','E','F'];
    const DESC_COL_START = 'G';
    const DESC_COL_END   = 'J';
    
    // 行の高さと、1ブロック（半ページ）あたりの使用行数
    const ROW_HEIGHT_PT = 24;   // 1行=24pt（見え方に応じて調整可）
    const BLOCK_ROWS    = 28;   // 画像＋説明で使用する行数（“半ページ”相当）
    
    // 列幅→ピクセル、行高(ポイント)→ピクセルの簡易換算
    const colPixels = (colIdx) => (ws.getColumn(colIdx+1).width || 10) * 7;
    const rowPixels = (rowIdx) => (ws.getRow(rowIdx).height || 18) * 1.333;
    const sumColPixels = (letters) =>
      letters.reduce((sum, L) => sum + colPixels(colLetterToIndex(L)), 0);
    
    // 画像領域の横幅（ピクセル）
    const containerW = sumColPixels(IMAGE_COLS);
    
    for (let i = 0; i < photos.length; i++) {
      const p = photos[i];
    
      // ブロックの行高を確保（“半ページ”分の高さを作る）
      for (let r = startRow; r < startRow + BLOCK_ROWS; r++) {
        ws.getRow(r).height = ROW_HEIGHT_PT;
      }
    
      // 画像領域の高さ（ピクセル）
      let containerH = 0;
      for (let rr = startRow; rr < startRow + BLOCK_ROWS; rr++) {
        containerH += rowPixels(rr);
      }
    
      // 画像実寸からフィットサイズを計算（縦横比維持）
      const { w: imgW, h: imgH } = await getImageDim(p.dataUrl);
      const ratioW = containerW / imgW;
      const ratioH = containerH / imgH;
      const ratio  = Math.min(ratioW, ratioH);
      const drawW  = Math.max(1, Math.floor(imgW * ratio));
      const drawH  = Math.max(1, Math.floor(imgH * ratio));
    
      // 左側（B列の位置）に画像を描画
      ws.addImage(
        wb.addImage({
          base64: p.dataUrl.split(',')[1],
          extension: p.dataUrl.startsWith('data:image/png') ? 'png' : 'jpeg',
        }),
        {
          tl:  { col: colLetterToIndex(IMAGE_COLS[0]), row: startRow - 1 },
          ext: { width: drawW, height: drawH },
        }
      );
    
      // 右側（G〜J列）に説明欄を縦いっぱいで作成
      const descRange = `${DESC_COL_START}${startRow}:${DESC_COL_END}${startRow + BLOCK_ROWS - 1}`;
      ws.mergeCells(descRange);
      const descCell = ws.getCell(`${DESC_COL_START}${startRow}`);
      descCell.value     = p.caption || '';
      descCell.alignment = { wrapText: true, vertical: 'top' };
      descCell.font      = { name: 'Meiryo UI' };
      descCell.border    = {
        top:    borderThin,
        left:   borderThin,
        bottom: borderThin,
        right:  borderThin,
      };
    
      // 次の画像ブロックへ（“半ページ”分の高さを使ったのでその分進める）
      startRow += (BLOCK_ROWS + 2); // +2は余白
    }
  }
  for (const e of entries) {
    await addOneEntrySheet(e);
  }

  // 目次
  let tocRow = 3;
  for (const e of entries) {
    wsToc.getCell(`A${tocRow}`).value = { text: e.displayLabel, hyperlink: `#'${e.sheetName}'!A1` };
    wsToc.getCell(`A${tocRow}`).style = linkStyle;
    const photosCount = Array.isArray(e.model.photos) ? e.model.photos.length : 0;
    wsToc.getCell(`B${tocRow}`).value = `写真 ${photosCount}枚`;
    wsToc.getCell(`B${tocRow}`).font  = { name: 'Meiryo UI' };
    tocRow++;
  }

  // 保存
  const buf = await wb.xlsx.writeBuffer();
  // 例: 「案件名_現地調査レポート.xlsx」
  // 空の場合は「現地調査レポート.xlsx」
  const title = (projectTitle || '').trim();
  const suffix = '現地調査レポート';
  const namePart = title ? `${title}_${suffix}` : suffix;
  
  // Windows等で無効な文字を避けるため簡易サニタイズ
  const safeNamePart = namePart.replace(/[\\/:*?"<>|]/g, '_');
  
  const fileName = `${safeNamePart}.xlsx`;
  window.saveAs(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fileName);
}



