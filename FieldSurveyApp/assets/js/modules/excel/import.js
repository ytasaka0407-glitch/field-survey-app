// FieldSurveyApp/assets/js/modules/excel/import.js
import { ensureSingle, ensureMulti, getOrInitStationData, selectedCategories, addSharedStation, sharedStations } from '../state.js';
import { toInputDateString, stationIdFromName } from '../utils.js';
import { getSchemaFor } from '../ui/schemas.js';

// 単一カテゴリで“基本欄”として上部に出力した項目（export.js と対応させる）
const CORE_SINGLE_EXTRA_FIELDS = [
  { key: 'method', label: '設置方法' }, // 追加項目を基本欄として扱う
];

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

    // 単一カテゴリの“基本欄追加”をA列ラベルで探索してB列の値を取得（追加項目の見出し「追加項目」より上）
    if (!isMulti && CORE_SINGLE_EXTRA_FIELDS.length) {
      for (const f of CORE_SINGLE_EXTRA_FIELDS) {
        // details が B5:J6 なので、その下（7行目以降）を「追加項目」見出しに当たるまで探索
        for (let r = 7; r < 200; r++) {
          const aText = String(ws.getCell(`A${r}`).value ?? '').trim();
          if (!aText) continue;
          if (aText === '追加項目') break; // ここより下は“追加項目”セクション
          if (aText === (f.label || f.key)) {
            const v = ws.getCell(`B${r}`).value;
            model[f.key] = (v ?? '').toString();
            break;
          }
        }
      }
    }

    // 追加項目の取り込み（“追加項目”見出しの位置を動的に検出）
    const schema = getSchemaFor(catName, isMulti ? 'multi' : 'single');

    // まず「追加項目」という見出し行を探す（A列にある）
    let extraHeaderRow = null;
    for (let r = 7; r < 200; r++) {
      const a = String(ws.getCell(`A${r}`).value ?? '').trim();
      if (a === '追加項目') { extraHeaderRow = r; break; }
    }

    if (extraHeaderRow != null) {
      let row = extraHeaderRow + 1;
      // 3連続空行で終了
      let emptyCount = 0;

      // “基本欄として扱うキー”は追加項目の候補から除外したいが、旧形式ファイルの互換のため、
      // ここでは禁止キー（date/location/details/photos）のみ除外し、methodは未設定の場合のみ取り込む
      const FORBIDDEN_KEYS = new Set(['date','location','details','photos']);

      while (row < extraHeaderRow + 1 + 200) {
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

        // 優先: K列のキー、なければラベル名からschemaで推測
        let targetKey = keyText;
        if (!targetKey) {
          const f = schema.find(s => (s.label || s.key) === labelText);
          targetKey = f?.key || '';
        }

        if (targetKey && !FORBIDDEN_KEYS.has(targetKey)) {
          const valStr = (valRaw ?? '').toString();
          // “method”は既に基本欄で取得済みなら上書きしない
          if (targetKey === 'method' && model[targetKey]) {
            // スキップ
          } else {
            model[targetKey] = valStr;
          }
        }
        row++;
      }
    }
  });
}
