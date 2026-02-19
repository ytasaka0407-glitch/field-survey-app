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

    const colWidths = [20, 16, 16, 18, 16, 16, 18, 16, 16, 18, 10, 10]; // A〜L
    colWidths.forEach((w, i) => ws.getColumn(i+1).width = w);
    ws.getColumn(11).hidden = true; // K列（フィールドキー格納）

    ws.getCell('A1').value = entry.displayLabel;
    ws.getCell('A1').style = sectionTitleStyle;

    // 基本項目（並び: 調査日, 設置場所, 新設/既設, 設置方法(新設時), OK/NG, NG理由(NG時), その他調査内容）
    // 調査日
    ws.getCell('A3').value = '調査日'; ws.getCell('A3').style = labelStyle;
    const d = fromInputDate(entry.model.date || projectDateStr);
    if (d) { ws.getCell('B3').value = d; ws.getCell('B3').numFmt = 'yyyy/mm/dd'; }
    else { ws.getCell('B3').value = '-'; }
    ws.getCell('B3').font = { name: 'Meiryo UI' };
    ws.getCell('B3').alignment = { horizontal: 'left' };

    // 設置場所
    ws.getCell('A4').value = '設置場所'; ws.getCell('A4').style = labelStyle;
    ws.mergeCells('B4:J4'); ws.getCell('B4').value = entry.model.location || '-';
    ws.getCell('B4').font  = { name: 'Meiryo UI' };

    // 以降の基本項目
    let nextRow = 5;

    // 新設/既設流用
    ws.getCell(`A${nextRow}`).value = '新設/既設流用';
    ws.getCell(`A${nextRow}`).style = labelStyle;
    ws.mergeCells(`B${nextRow}:J${nextRow}`);
    ws.getCell(`B${nextRow}`).value = (entry.model.installType === 'reuse') ? '既設流用' : '新設';
    ws.getCell(`B${nextRow}`).font = { name: 'Meiryo UI' };
    nextRow++;

    // 設置方法（新設時のみ）
    if (entry.model.installType === 'new') {
      ws.getCell(`A${nextRow}`).value = '設置方法';
      ws.getCell(`A${nextRow}`).style = labelStyle;
      ws.mergeCells(`B${nextRow}:J${nextRow}`);
      ws.getCell(`B${nextRow}`).value = (entry.model.method ?? '').toString() || '-';
      ws.getCell(`B${nextRow}`).font = { name: 'Meiryo UI' };
    
      // ↓ 追加：設置方法の行の下に罫線
      ws.getCell(`A${nextRow}`).border = { bottom: borderThin };
      ws.getCell(`B${nextRow}`).border = { bottom: borderThin };
    
      nextRow++;
    }

    // 系統図との整合性（OK/NG）
    ws.getCell(`A${nextRow}`).value = '系統図との整合性';
    ws.getCell(`A${nextRow}`).style = labelStyle;
    ws.mergeCells(`B${nextRow}:J${nextRow}`);
    ws.getCell(`B${nextRow}`).value = (entry.model.diagramStatus === 'ng') ? 'NG' : 'OK';
    ws.getCell(`B${nextRow}`).font = { name: 'Meiryo UI' };
    nextRow++;

    // NG理由（NG時のみ）
    if (entry.model.diagramStatus === 'ng') {
      ws.getCell(`A${nextRow}`).value = 'NG理由';
      ws.getCell(`A${nextRow}`).style = labelStyle;
      ws.mergeCells(`B${nextRow}:J${nextRow}`);
      ws.getCell(`B${nextRow}`).value = (entry.model.diagramNgReason ?? '').toString() || '-';
      ws.getCell(`B${nextRow}`).alignment = { wrapText: true, vertical: 'top' };
      ws.getCell(`B${nextRow}`).font = { name: 'Meiryo UI' };
    
      // ↓ 追加：NG理由の行の下に罫線
      ws.getCell(`A${nextRow}`).border = { bottom: borderThin };
      ws.getCell(`B${nextRow}`).border = { bottom: borderThin };
    
      nextRow++;
    }

    // その他調査内容（1〜2行分の高さを確保）
    ws.getCell(`A${nextRow}`).value = 'その他調査内容';
    ws.getCell(`A${nextRow}`).style = labelStyle;
    ws.mergeCells(`B${nextRow}:J${nextRow+1}`);
    ws.getCell(`B${nextRow}`).value     = (entry.model.details || '').replace(/\r?\n/g, '\n');
    ws.getCell(`B${nextRow}`).alignment = { wrapText: true, vertical: 'top' };
    ws.getCell(`B${nextRow}`).font      = { name: 'Meiryo UI' };
    ws.getRow(nextRow).height = 22;
    ws.getRow(nextRow+1).height = 22;
    nextRow += 2;

    // 区切り線（任意）
    ['A3','A4',`A${nextRow-2}`].forEach(addr => ws.getCell(addr).border = { bottom: borderThin });
    ['B3','B4',`B${nextRow-2}`].forEach(addr => ws.getCell(addr).border = { bottom: borderThin });

    // 追加項目（スキーマ定義に基づく。コア扱いのキーは除外）
    const schema = getSchemaFor(entry.cat, entry.type === 'multi' ? 'multi' : 'single');
    const CORE_KEYS = new Set(['date','location','installType','method','diagramStatus','diagramNgReason','details','photos']);
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

    // 写真セクション（左右分割）
    let startRow = nextRow + 1;
    const photos = Array.isArray(entry.model.photos) ? entry.model.photos : [];

    const IMAGE_COLS = ['B','C','D','E','F'];
    const DESC_COL_START = 'G';
    const DESC_COL_END   = 'J';

    const ROW_HEIGHT_PT = 24;
    const BLOCK_ROWS    = 11;

    const colLetterToIndex = (L) => L.charCodeAt(0) - 65;
    const colPixels = (colIdx) => (ws.getColumn(colIdx+1).width || 10) * 7;
    const rowPixels = (rowIdx) => (ws.getRow(rowIdx).height || 18) * 1.333;
    const sumColPixels = (letters) =>
      letters.reduce((sum, L) => sum + colPixels(colLetterToIndex(L)), 0);

    const containerW = sumColPixels(IMAGE_COLS);

    for (let i = 0; i < photos.length; i++) {
      const p = photos[i];

      for (let r = startRow; r < startRow + BLOCK_ROWS; r++) {
        ws.getRow(r).height = ROW_HEIGHT_PT;
      }

      let containerH = 0;
      for (let rr = startRow; rr < startRow + BLOCK_ROWS; rr++) {
        containerH += rowPixels(rr);
      }

      const { w: imgW, h: imgH } = await getImageDim(p.dataUrl);
      const ratioW = containerW / imgW;
      const ratioH = containerH / imgH;
      const ratio  = Math.min(ratioW, ratioH);
      const drawW  = Math.max(1, Math.floor(imgW * ratio));
      const drawH  = Math.max(1, Math.floor(imgH * ratio));

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

      const descRange = `${DESC_COL_START}${startRow}:${DESC_COL_END}${startRow + BLOCK_ROWS - 1}`;
      ws.mergeCells(descRange);
      const descCell = ws.getCell(`${DESC_COL_START}${startRow}`);
      descCell.value = (p.caption || '').replace(/\r\n?/g, '\n');
      descCell.alignment = { wrapText: true, vertical: 'top' };
      descCell.font      = { name: 'Meiryo UI' };
      descCell.border    = {
        top:    borderThin,
        left:   borderThin,
        bottom: borderThin,
        right:  borderThin,
      };

      startRow += (BLOCK_ROWS + 2);
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

  const buf = await wb.xlsx.writeBuffer();
  const title = (projectTitle || '').trim();
  const suffix = '現地調査レポート';
  const namePart = title ? `${title}_${suffix}` : suffix;
  const safeNamePart = namePart.replace(/[\\/:*?"<>|]/g, '_');
  const fileName = `${safeNamePart}.xlsx`;
  window.saveAs(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fileName);
}

