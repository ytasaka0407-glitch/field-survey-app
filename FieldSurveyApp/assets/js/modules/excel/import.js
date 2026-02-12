// FieldSurveyApp/assets/js/modules/excel/import.js
import { ensureSingle, ensureMulti, getOrInitStationData, selectedCategories, addSharedStation, sharedStations } from '../state.js';
import { toInputDateString, stationIdFromName } from '../utils.js';
import { getSchemaFor } from '../ui/schemas.js';

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

  // カテゴリシート
  wb.worksheets.forEach((ws) => {
    if (!ws) return;
    const name = ws.name;
    if (name === '表紙' || name === '目次') return;

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

    // 追加項目の取り込み
    const schema = getSchemaFor(catName, isMulti ? 'multi' : 'single');
    const extraFieldKeys = schema
      .filter(f => !['date','location','details','photos'].includes(f.key))
      .map(f => f.key);

    // A7が「追加項目」ならその下を読み取る
    const marker = String(ws.getCell('A7').value ?? '').trim();
    if (marker === '追加項目') {
      let row = 8;
      // 200行程度を上限に読み取り（空行が続いたら早期終了）
      let emptyCount = 0;
      while (row < 208) {
        const labelCell = ws.getCell(`A${row}`);
        const valueCell = ws.getCell(`B${row}`);
        const keyCell   = ws.getCell(`K${row}`); // エクスポート時に埋めたキー（K列は非表示）

        const labelText = String(labelCell.value ?? '').trim();
        const valRaw    = valueCell.value;
        const keyText   = String(keyCell.value ?? '').trim();

        if (!labelText && !valRaw && !keyText) {
          emptyCount++;
          if (emptyCount >= 3) break; // 3連続で空なら終わり
          row++;
          continue;
        }
        emptyCount = 0;

        let targetKey = keyText;
        // キーが空ならラベル名から推測（重複ラベルを避けるためキーを使うことを推奨）
        if (!targetKey) {
          // ラベル一致でschemaから探す
          const f = schema.find(s => (s.label || s.key) === labelText);
          targetKey = f?.key || '';
        }
        if (targetKey && extraFieldKeys.includes(targetKey)) {
          const valStr = (valRaw ?? '').toString();
          model[targetKey] = valStr;
        }
        row++;
      }
    }
  });
}
