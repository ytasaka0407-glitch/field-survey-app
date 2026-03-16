import { initDefaults, dataMap, selectedCategories, sharedStations, setSharedStations, setProjectDatePrev, ensureSingle, ensureMulti, getOrInitStationData, resetAllState } from './modules/state.js';
import { initCategoryUI } from './modules/ui/categories.js';
import { initBlocksUI } from './modules/ui/blocks.js';
import { saveDraft, loadDraft, migrateSharedStationsFromLegacy, hydratePhotosFromIDB } from './modules/storage.js';
import { exportToExcel } from './modules/excel/export.js';
import { importFromExcel } from './modules/excel/import.js';

export function bootstrapApp() {
  // DOM参照
  const projectTitleEl = document.getElementById("projectTitle");
  const projectDateEl  = document.getElementById("projectDate");
  const categorySelectArea = document.getElementById("categorySelectArea");
  const newSingleCategoryInput = document.getElementById("newSingleCategoryInput");
  const addSingleCategoryBtn = document.getElementById("addSingleCategoryBtn");
  const newMultiCategoryInput = document.getElementById("newMultiCategoryInput");
  const addMultiCategoryBtn = document.getElementById("addMultiCategoryBtn");
  const singleCatsContainer = document.getElementById("singleCatsContainer");
  const multiStationsContainer = document.getElementById("multiStationsContainer");
  const clearBtn = document.getElementById('clearBtn');
  const saveBtn = document.getElementById("saveBtn");
  const loadBtn = document.getElementById("loadBtn");
  const exportExcelBtn = document.getElementById("exportExcelBtn");
  const importExcelBtn  = document.getElementById('importExcelBtn');
  const importExcelFile = document.getElementById('importExcelFile');
  const newStationInput = document.getElementById("newStationInput");
  const addStationBtn = document.getElementById("addStationBtn");
  const stationListEl = document.getElementById("stationList");

  // 初期化
  initDefaults();

  // UI初期化
  const catUI = initCategoryUI(
    {
      categorySelectArea,
      newSingleCategoryInput, addSingleCategoryBtn,
      newMultiCategoryInput, addMultiCategoryBtn,
      stationListEl, newStationInput, addStationBtn,
      projectDateEl
    },
    {
      onSelectionChanged: () => blocks.renderCategories(),
      onStationsChanged: () => blocks.renderCategories(),
    }
  );
  const blocks = initBlocksUI({ singleCatsContainer, multiStationsContainer, projectDateEl });

  // 表紙の調査日自動反映（個別上書きは保持）
  let projectDatePrev = "";
  projectDateEl.addEventListener("change", (e) => {
    const newDate = e.target.value || "";
    Object.keys(dataMap).forEach(cat => {
      const mode = dataMap[cat]?.mode || "single";
      if (mode === "single") {
        const v = ensureSingle(cat, projectDateEl.value || "");
        if (!v.date || v.date === projectDatePrev) {
          v.date = newDate;
          const idBase = cat; // blocks側で再描画するためDOM直接操作は最小化
        }
      } else {
        ensureMulti(cat, projectDateEl.value || "");
        const sd = dataMap[cat].stationData || {};
        Object.keys(sd).forEach(stId => {
          const s = sd[stId];
          if (!s.date || s.date === projectDatePrev) {
            s.date = newDate;
          }
        });
      }
    });
    projectDatePrev = newDate;
    setProjectDatePrev(newDate);
    blocks.renderCategories();
  });

  // 画面初期描画
  catUI.renderCategorySelector();
  blocks.renderCategories();

  // 保存/読込
  saveBtn.addEventListener("click", async () => {
    await saveDraft(projectTitleEl.value || "", projectDateEl.value || "");
    alert("下書きを保存しました");
  });
  loadBtn.addEventListener("click", async () => {
    const payload = loadDraft();
    if (!payload) { alert("保存された下書きがありません"); return; }
    try {
      projectTitleEl.value = payload.projectTitle || "";
      projectDateEl.value  = payload.projectDate || "";
      setProjectDatePrev(payload.projectDate || "");
  
      Object.keys(dataMap).forEach((k) => delete dataMap[k]);
      Object.assign(dataMap, payload.data || {});
      if (payload.sharedStations && Array.isArray(payload.sharedStations)) {
        setSharedStations(payload.sharedStations.slice());
      } else {
        migrateSharedStationsFromLegacy(projectDateEl.value || "");
      }
      selectedCategories.clear();
      (payload.selected || []).forEach((c) => selectedCategories.add(c));
  
      // 写真のdataUrlをIndexedDBから復元
      await hydratePhotosFromIDB();
  
      catUI.renderCategorySelector();
      blocks.renderCategories();
      alert("下書きを読み込みました");
    } catch (e) {
      console.error(e);
      alert("読み込みに失敗しました");
    }
  });

  // Excel出力/入力
  exportExcelBtn.addEventListener('click', async () => {
    await exportToExcel((projectTitleEl.value || '').trim(), projectDateEl.value || '', sharedStations);
  });
  importExcelBtn?.addEventListener('click', () => importExcelFile.click());
  importExcelFile?.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      await importFromExcel(file, projectTitleEl, projectDateEl, setProjectDatePrev);
      catUI.renderCategorySelector();
      blocks.renderCategories();
      alert('Excelを読み込みました');
      importExcelFile.value = '';
    } catch (err) {
      console.error(err);
      alert('Excelの読み込みに失敗しました。ファイル形式とフォーマットをご確認ください。');
    }
  });

  // 画面クリア（iPhone対策：touchendも拾う。二重発火はガード）
  const doClear = () => {
    const ok = window.confirm('現在表示されている内容が全て消えますが、よろしいですか？');
    if (!ok) return;

    // 表紙入力もクリア
    projectTitleEl.value = '';
    projectDateEl.value = '';
    setProjectDatePrev('');

    // stateクリア → 初期化
    resetAllState();

    // 再描画
    catUI.renderCategorySelector();
    blocks.renderCategories();
  };
  
  let clearFiring = false;

  const fireClear = (e) => {
    if (e) e.preventDefault();
    if (clearFiring) return;
    clearFiring = true;
    try {
      doClear();
    } finally {
      // 同一ジェスチャで click/touch/pointer が連続発火しても1回にする
      setTimeout(() => { clearFiring = false; }, 400);
    }
  };

  // PC/Android はこちらが効くことが多い
  clearBtn?.addEventListener('pointerup', fireClear);

  // iPhone(iOS Safari) の保険
  clearBtn?.addEventListener('touchend', fireClear, { passive: false });

  // さらに最後の保険（pointer/touchが効かない環境向け）
  clearBtn?.addEventListener('click', fireClear);
  
}

