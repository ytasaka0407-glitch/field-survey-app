import { dataMap, selectedCategories, ensureSingle, ensureMulti, getOrInitStationData } from '../state.js';
import { escapeHtml, hashId, readFileAsDataURL, resizeImage } from '../utils.js';

export function initBlocksUI(refs) {
  const { singleCatsContainer, multiStationsContainer, projectDateEl, sharedStations } = refs;

  function renderPhotoList(photoArr, listEl) {
    listEl.innerHTML = "";
    photoArr.forEach((p, idx) => {
      const item = document.createElement("div");
      item.className = "photo-item";
      item.innerHTML = `
        <img src="${p.dataUrl}" alt="photo ${idx + 1}">
        <div class="caption">
          <input type="text" placeholder="説明（キャプション）" value="${p.caption || ""}" />
        </div>
        <div class="footer">
          <span>${p.name || `photo_${idx + 1}.jpg`}</span>
          <button class="remove">削除</button>
        </div>
      `;
      item.querySelector("input").addEventListener("input", (e) => { p.caption = e.target.value || ""; });
      item.querySelector(".remove").addEventListener("click", () => {
        photoArr.splice(idx, 1);
        renderPhotoList(photoArr, listEl);
      });
      listEl.appendChild(item);
    });
  }

  function handleFilesForSingle(cat, fileList) {
    const v = ensureSingle(cat, projectDateEl.value || "");
    (async () => {
      for (const file of fileList) {
        if (!file.type.startsWith("image/")) continue;
        const dataUrl = await readFileAsDataURL(file);
        const resized = await resizeImage(dataUrl, 1280);
        v.photos.push({ dataUrl: resized, name: file.name, caption: "" });
      }
      const idBase = hashId(cat);
      const listEl = document.getElementById("list_" + idBase);
      if (listEl) renderPhotoList(v.photos, listEl);
    })();
  }

  function handleFilesForStationCategory(cat, stationId, fileList, segEl) {
    const stData = getOrInitStationData(cat, stationId, projectDateEl.value || "");
    (async () => {
      for (const file of fileList) {
        if (!file.type.startsWith("image/")) continue;
        const dataUrl = await readFileAsDataURL(file);
        const resized = await resizeImage(dataUrl, 1280);
        stData.photos.push({ dataUrl: resized, name: file.name, caption: "" });
      }
      const idBase = hashId(cat + "|" + stationId);
      const listEl = segEl.querySelector("#list_" + idBase);
      if (listEl) renderPhotoList(stData.photos, listEl);
    })();
  }

  function buildSingleCategoryBlock(cat) {
    ensureSingle(cat, projectDateEl.value || "");
    const idBase = hashId(cat);
    const v = dataMap[cat];
    const block = document.createElement("div");
    block.className = "card";
    block.innerHTML = `
      <h3>${escapeHtml(cat)} </h3>
      <div class="form-row">
        <label for="date_${idBase}">調査日</label>
        <input id="date_${idBase}" type="date" value="${v.date || ""}" />
      </div>
      <div class="form-row">
        <label for="loc_${idBase}">設置場所</label>
        <input id="loc_${idBase}" type="text" placeholder="設置する箇所をなるべく詳細に記載" value="${v.location || ""}" />
      </div>
      <div class="form-row">
        <label for="det_${idBase}">調査内容</label>
        <textarea id="det_${idBase}" placeholder="現地での調査事項、寸法、注意点など">${v.details || ""}</textarea>
      </div>
      <div class="form-row">
        <label>写真</label>
        <div>
          <div class="buttons">
            <input id="cam_${idBase}" type="file" accept="image/*" capture="environment" multiple style="display:none" />
            <button id="camBtn_${idBase}" class="btn primary">撮影する（カメラ）</button>
            <input id="file_${idBase}" type="file" accept="image/*" multiple style="display:none" />
            <button id="fileBtn_${idBase}" class="btn">画像を選択</button>
          </div>
          <div class="photos">
            <div id="list_${idBase}" class="photo-list"></div>
          </div>
        </div>
      </div>
    `;
    block.querySelector("#date_" + idBase).addEventListener("change", (e) => (v.date = e.target.value));
    block.querySelector("#loc_" + idBase).addEventListener("input", (e) => (v.location = e.target.value));
    block.querySelector("#det_" + idBase).addEventListener("input", (e) => (v.details = e.target.value));

    const camInput = block.querySelector("#cam_" + idBase);
    const fileInput = block.querySelector("#file_" + idBase);
    block.querySelector("#camBtn_" + idBase).addEventListener("click", () => camInput.click());
    block.querySelector("#fileBtn_" + idBase).addEventListener("click", () => fileInput.click());

    camInput.addEventListener("change", (e) => handleFilesForSingle(cat, e.target.files));
    fileInput.addEventListener("change", (e) => handleFilesForSingle(cat, e.target.files));

    renderPhotoList(v.photos, block.querySelector("#list_" + idBase));
    return block;
  }

  function buildStationBlock(st, multiCats) {
    const wrap = document.createElement("div");
    wrap.className = "station-block";
    wrap.innerHTML = `
      <div class="station-header">
        <div class="station-title">${escapeHtml(st.name)}</div>
      </div>
    `;
    if (!multiCats.length) {
      const p = document.createElement("p");
      p.className = "muted";
      p.textContent = "基地局項目が未選択です。上の項目選択から基地局項目を選んでください。";
      wrap.appendChild(p);
      return wrap;
    }
    multiCats.forEach(cat => {
      ensureMulti(cat, projectDateEl.value || "");
      const stData = getOrInitStationData(cat, st.id, projectDateEl.value || "");
      const idBase = hashId(cat + "|" + st.id);
      const seg = document.createElement("div");
      seg.className = "card";
      seg.style.margin = "8px 0";
      seg.innerHTML = `
        <div class="subsection-title">${escapeHtml(cat)} </div>
        <div class="form-row">
          <label for="date_${idBase}">調査日</label>
          <input id="date_${idBase}" type="date" value="${stData.date || ""}" />
        </div>
        <div class="form-row">
          <label for="loc_${idBase}">設置場所</label>
          <input id="loc_${idBase}" type="text" placeholder="設置する箇所をなるべく詳細に記載" value="${stData.location || ""}" />
        </div>
        <div class="form-row">
          <label for="det_${idBase}">調査内容</label>
          <textarea id="det_${idBase}" placeholder="現地での調査事項、寸法、注意点など">${stData.details || ""}</textarea>
        </div>
        <div class="form-row">
          <label>写真</label>
          <div>
            <div class="buttons">
              <input id="cam_${idBase}" type="file" accept="image/*" capture="environment" multiple style="display:none" />
              <button id="camBtn_${idBase}" class="btn primary">撮影する（カメラ）</button>
              <input id="file_${idBase}" type="file" accept="image/*" multiple style="display:none" />
              <button id="fileBtn_${idBase}" class="btn">画像を選択</button>
            </div>
            <div class="photos">
              <div id="list_${idBase}" class="photo-list"></div>
            </div>
          </div>
        </div>
      `;
      seg.querySelector("#date_" + idBase).addEventListener("change", (e) => (stData.date = e.target.value));
      seg.querySelector("#loc_" + idBase).addEventListener("input", (e) => (stData.location = e.target.value));
      seg.querySelector("#det_" + idBase).addEventListener("input", (e) => (stData.details = e.target.value));
      const camInput = seg.querySelector("#cam_" + idBase);
      const fileInput = seg.querySelector("#file_" + idBase);
      seg.querySelector("#camBtn_" + idBase).addEventListener("click", () => camInput.click());
      seg.querySelector("#fileBtn_" + idBase).addEventListener("click", () => fileInput.click());
      camInput.addEventListener("change", (e) => handleFilesForStationCategory(cat, st.id, e.target.files, seg));
      fileInput.addEventListener("change", (e) => handleFilesForStationCategory(cat, st.id, e.target.files, seg));
      renderPhotoList(stData.photos, seg.querySelector("#list_" + idBase));
      wrap.appendChild(seg);
    });
    return wrap;
  }

  function renderCategories() {
    // 指令センター
    singleCatsContainer.innerHTML = "";
    const selectedSingles = [...selectedCategories].filter(c => (dataMap[c]?.mode || "single") === "single").sort((a,b)=>a.localeCompare(b,"ja"));
    selectedSingles.forEach(cat => singleCatsContainer.appendChild(buildSingleCategoryBlock(cat)));
    // 基地局
    multiStationsContainer.innerHTML = "";
    const selectedMultis = [...selectedCategories].filter(c => (dataMap[c]?.mode) === "multi").sort((a,b)=>a.localeCompare(b,"ja"));
    if (!selectedMultis.length) {
      const p = document.createElement("p"); p.className = "muted";
      p.textContent = "基地局項目が未選択です。項目選択で基地局項目を選んでください。";
      multiStationsContainer.appendChild(p);
    } else if (!sharedStations.length) {
      const p2 = document.createElement("p"); p2.className = "muted";
      p2.textContent = "基地局が未登録です。項目選択で基地局を追加してください。";
      multiStationsContainer.appendChild(p2);
    } else {
      sharedStations.forEach(st => {
        selectedMultis.forEach(cat => ensureMulti(cat, projectDateEl.value || ""));
        multiStationsContainer.appendChild(buildStationBlock(st, selectedMultis));
      });
    }
  }

  return { renderCategories };
}