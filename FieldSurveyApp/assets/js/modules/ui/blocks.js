// FieldSurveyApp/assets/js/modules/ui/blocks.js
import { dataMap, selectedCategories, ensureSingle, ensureMulti, getOrInitStationData, sharedStations } from '../state.js';
import { escapeHtml, hashId } from '../utils.js';
import { getSchemaFor } from './schemas.js';
import { renderField } from './fields.js';

export function initBlocksUI(refs) {
  const { singleCatsContainer, multiStationsContainer, projectDateEl } = refs;

  function buildSingleCategoryBlock(cat) {
    const v = ensureSingle(cat, projectDateEl.value || '');
    const idBase = hashId(cat);
    const schema = getSchemaFor(cat, 'single');
    const block = document.createElement('div');
    block.className = 'card';
    const title = document.createElement('h3');
    title.textContent = cat; // ← escapeHtmlは不要（textContentは安全）
    block.appendChild(title);
    const inner = document.createElement('div');
    block.appendChild(inner);
    schema.forEach(field => {
      const fid = `${field.key}_${idBase}`;
      renderField(inner, v, field, fid);
    });
    return block;
  }

  function buildStationBlock(st, multiCats) {
    const wrap = document.createElement('div');
    wrap.className = 'station-block';
    // ここは innerHTML なので escapeHtml を使用
    wrap.innerHTML = `<div class="station-header"><div class="station-title">${escapeHtml(st.name)}</div></div>`;
    if (!multiCats.length) {
      const p = document.createElement('p');
      p.className = 'muted';
      p.textContent = '基地局項目が未選択です。上の項目選択から基地局項目を選んでください。';
      wrap.appendChild(p);
      return wrap;
    }
    multiCats.forEach(cat => {
      ensureMulti(cat, projectDateEl.value || '');
      const stData = getOrInitStationData(cat, st.id, projectDateEl.value || '');
      const idBase = hashId(cat + '|' + st.id);
      const schema = getSchemaFor(cat, 'multi');
      const seg = document.createElement('div');
      seg.className = 'card';
      seg.style.margin = '8px 0';
      const title = document.createElement('div');
      title.className = 'subsection-title';
      title.textContent = cat; // ← escapeHtmlは不要（textContent）
      seg.appendChild(title);
      const inner = document.createElement('div');
      seg.appendChild(inner);
      schema.forEach(field => {
        const fid = `${field.key}_${idBase}`;
        renderField(inner, stData, field, fid);
      });
      wrap.appendChild(seg);
    });
    return wrap;
  }

  function renderCategories() {
    // 指令センター
    singleCatsContainer.innerHTML = '';
    const selectedSingles = [...selectedCategories]
      .filter(c => (dataMap[c]?.mode || 'single') === 'single')
      .sort((a,b)=>a.localeCompare(b,'ja'));
    selectedSingles.forEach(cat => singleCatsContainer.appendChild(buildSingleCategoryBlock(cat)));

    // 基地局（state.sharedStations を常に最新で参照）
    multiStationsContainer.innerHTML = '';
    const selectedMultis = [...selectedCategories]
      .filter(c => (dataMap[c]?.mode) === 'multi')
      .sort((a,b)=>a.localeCompare(b,'ja'));
    if (!selectedMultis.length) {
      const p = document.createElement('p'); p.className = 'muted';
      p.textContent = '基地局項目が未選択です。項目選択で基地局項目を選んでください。';
      multiStationsContainer.appendChild(p);
    } else if (!sharedStations.length) {
      const p2 = document.createElement('p'); p2.className = 'muted';
      p2.textContent = '基地局が未登録です。項目選択で基地局を追加してください。';
      multiStationsContainer.appendChild(p2);
    } else {
      sharedStations.forEach(st => {
        multiStationsContainer.appendChild(buildStationBlock(st, selectedMultis));
      });
    }
  }

  return { renderCategories };
}
