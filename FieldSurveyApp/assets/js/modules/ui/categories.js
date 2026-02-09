import { dataMap, selectedCategories, sharedStations, addSharedStation, removeSharedStationById, ensureSingle, ensureMulti, defaultSingleCategories, defaultMultiCategories } from '../state.js';
import { escapeHtml, hashId, stationIdFromName } from '../utils.js';

export function initCategoryUI(refs, callbacks) {
  const {
    categorySelectArea,
    newSingleCategoryInput, addSingleCategoryBtn,
    newMultiCategoryInput, addMultiCategoryBtn,
    stationListEl, newStationInput, addStationBtn,
    projectDateEl
  } = refs;
  const { onSelectionChanged, onStationsChanged } = callbacks;

  function renderSharedStationsManager() {
    stationListEl.innerHTML = "";
    if (!sharedStations.length) {
      const empty = document.createElement("div");
      empty.className = "muted";
      empty.textContent = "基地局は未登録です。名前を入力して「基地局を追加」を押してください。";
      stationListEl.appendChild(empty);
      return;
    }
    sharedStations.forEach(st => {
      const chip = document.createElement("div");
      chip.className = "station-chip";
      chip.innerHTML = `
        <span class="name">${escapeHtml(st.name)}</span>
        <button class="remove" title="削除">×</button>
      `;
      chip.querySelector(".remove").addEventListener("click", () => {
        if (confirm(`基地局「${st.name}」を削除します。全ての基地局項目からも削除されます。よろしいですか？`)) {
          removeSharedStationById(st.id);
          renderCategorySelector();
          onStationsChanged();
        }
      });
      stationListEl.appendChild(chip);
    });
  }

  function renderCategorySelector() {
    const singles = [];
    const multis = [];
    Object.keys(dataMap).forEach((cat) => {
      const m = dataMap[cat]?.mode || "single";
      if (m === "multi") multis.push(cat);
      else singles.push(cat);
    });
    singles.sort((a, b) => a.localeCompare(b, "ja"));
    multis.sort((a, b) => a.localeCompare(b, "ja"));

    categorySelectArea.innerHTML = "";

    const singleCol = document.createElement("div");
    const multiCol = document.createElement("div");
    singleCol.innerHTML = `<h3 style="margin-bottom:8px;">指令センター項目</h3>`;
    multiCol.innerHTML = `<h3 style="margin-bottom:8px;">基地局項目</h3>`;

    const singleList = document.createElement("div"); singleList.className = "category-list";
    const multiList = document.createElement("div");  multiList.className = "category-list";

    singles.forEach((cat) => {
      const id = "cat_" + hashId("single|" + cat);
      const wrapper = document.createElement("div");
      wrapper.className = "category-item";
      wrapper.innerHTML = `
        <div class="category-left">
          <input type="checkbox" id="${id}" ${selectedCategories.has(cat) ? "checked" : ""}>
          <span>${escapeHtml(cat)}</span>
        </div>
        <div class="category-actions">
          ${defaultSingleCategories.includes(cat) ? "" : `<button class="icon-btn" title="削除">🗑️</button>`}
        </div>
      `;
      wrapper.querySelector("input").addEventListener("change", (e) => {
        if (e.target.checked) {
          ensureSingle(cat, projectDateEl.value || "");
          selectedCategories.add(cat);
        } else {
          selectedCategories.delete(cat);
        }
        onSelectionChanged();
      });
      const delBtn = wrapper.querySelector(".icon-btn");
      if (delBtn) {
        delBtn.addEventListener("click", () => {
          if (confirm(`項目「${cat}」を削除しますか？（入力済みデータも削除されます）`)) {
            selectedCategories.delete(cat);
            delete dataMap[cat];
            renderCategorySelector();
            onSelectionChanged();
          }
        });
      }
      singleList.appendChild(wrapper);
    });

    multis.forEach((cat) => {
      const id = "cat_" + hashId("multi|" + cat);
      const wrapper = document.createElement("div");
      wrapper.className = "category-item";
      wrapper.innerHTML = `
        <div class="category-left">
          <input type="checkbox" id="${id}" ${selectedCategories.has(cat) ? "checked" : ""}>
          <span>${escapeHtml(cat)}</span>
        </div>
        <div class="category-actions">
          ${defaultMultiCategories.includes(cat) ? "" : `<button class="icon-btn" title="削除">🗑️</button>`}
        </div>
      `;
      wrapper.querySelector("input").addEventListener("change", (e) => {
        if (e.target.checked) {
          ensureMulti(cat, projectDateEl.value || "");
          selectedCategories.add(cat);
        } else {
          selectedCategories.delete(cat);
        }
        onSelectionChanged();
      });
      const delBtn = wrapper.querySelector(".icon-btn");
      if (delBtn) {
        delBtn.addEventListener("click", () => {
          if (confirm(`項目「${cat}」を削除しますか？（入力済みデータも削除されます）`)) {
            selectedCategories.delete(cat);
            delete dataMap[cat];
            renderCategorySelector();
            onSelectionChanged();
          }
        });
      }
      multiList.appendChild(wrapper);
    });

    singleCol.appendChild(singleList);
    multiCol.appendChild(multiList);
    categorySelectArea.appendChild(singleCol);
    categorySelectArea.appendChild(multiCol);

    renderSharedStationsManager();
  }

  function bindEnterToButton(inputEl, buttonEl) {
    if (!inputEl || !buttonEl) return;
    inputEl.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); buttonEl.click(); }
    });
  }

  function addSharedStationByName(nm) {
    const name = (nm || "").trim();
    if (!name) return;
    const id = stationIdFromName(name);
    if (sharedStations.some(s => s.id === id)) {
      alert("同名の基地局が既に存在します");
      return;
    }
    addSharedStation({ id, name });
    Object.keys(dataMap).forEach(cat => {
      if ((dataMap[cat]?.mode) === "multi") {
        ensureMulti(cat, projectDateEl.value || "");
      }
    });
    renderCategorySelector();
    onStationsChanged();
  }

  addSingleCategoryBtn.addEventListener("click", () => {
    const name = (newSingleCategoryInput.value || "").trim();
    if (!name) return;
    ensureSingle(name, projectDateEl.value || "");
    selectedCategories.add(name);
    newSingleCategoryInput.value = "";
    renderCategorySelector(); onSelectionChanged();
  });
  addMultiCategoryBtn.addEventListener("click", () => {
    const name = (newMultiCategoryInput.value || "").trim();
    if (!name) return;
    ensureMulti(name, projectDateEl.value || "");
    selectedCategories.add(name);
    newMultiCategoryInput.value = "";
    renderCategorySelector(); onSelectionChanged();
  });
  addStationBtn.addEventListener("click", () => {
    const nm = newStationInput.value.trim();
    if (!nm) return;
    addSharedStationByName(nm);
    newStationInput.value = "";
  });

  bindEnterToButton(newSingleCategoryInput, addSingleCategoryBtn);
  bindEnterToButton(newMultiCategoryInput, addMultiCategoryBtn);
  bindEnterToButton(newStationInput, addStationBtn);

  return { renderCategorySelector };
}