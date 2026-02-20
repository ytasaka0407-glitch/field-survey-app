import { dataMap, selectedCategories, sharedStations, setSharedStations } from './state.js';
import { stationIdFromName } from './utils.js';
import { multiCategoryDefaults } from './state.js';
import { putPhotoDataUrl, getPhotoDataUrl } from './idb-photos.js';

export async function saveDraft(projectTitle, projectDate) {
  // 写真データを IndexedDB に退避し、payload には参照IDのみ残す
  const payload = {
    projectTitle: projectTitle || "",
    projectDate: projectDate || "",
    selected: [...selectedCategories],
    data: {},
    sharedStations,
  };

  // 深拷貝しつつ dataUrl をID参照化
  const clone = JSON.parse(JSON.stringify(dataMap));

  // 1) single
  for (const cat of Object.keys(clone)) {
    const v = clone[cat];
    if (!v) continue;
    if ((v.mode || 'single') === 'single') {
      if (Array.isArray(v.photos)) {
        for (const p of v.photos) {
          if (!p) continue;
          if (!p.id) p.id = 'ph_' + Math.random().toString(36).slice(2);
          if (p.dataUrl) {
            try { await putPhotoDataUrl(p.id, p.dataUrl); } catch {}
            p.dataUrl = ''; // localStorage には持たせない
          }
        }
      }
      payload.data[cat] = v;
    } else if (v.mode === 'multi') {
      const stationData = v.stationData || {};
      for (const stId of Object.keys(stationData)) {
        const s = stationData[stId];
        if (Array.isArray(s.photos)) {
          for (const p of s.photos) {
            if (!p) continue;
            if (!p.id) p.id = 'ph_' + Math.random().toString(36).slice(2);
            if (p.dataUrl) {
              try { await putPhotoDataUrl(p.id, p.dataUrl); } catch {}
              p.dataUrl = '';
            }
          }
        }
      }
      payload.data[cat] = { mode: 'multi', stationData };
    }
  }

  localStorage.setItem("surveyDraftV4", JSON.stringify(payload));
}

export function loadDraft() {
  const raw = localStorage.getItem("surveyDraftV4")
    || localStorage.getItem("surveyDraftV3")
    || localStorage.getItem("surveyDraftV2");
  if (!raw) return null;
  try {
    const payload = JSON.parse(raw);
    return payload;
  } catch {
    return null;
  }
}

// IndexedDB に退避している写真をメモリ(dataMap)へ復元
export async function hydratePhotosFromIDB() {
  // single
  for (const cat of Object.keys(dataMap)) {
    const v = dataMap[cat];
    if (!v) continue;
    if ((v.mode || 'single') === 'single') {
      if (Array.isArray(v.photos)) {
        for (const p of v.photos) {
          if (p && !p.dataUrl && p.id) {
            try { p.dataUrl = await getPhotoDataUrl(p.id); } catch {}
          }
        }
      }
    } else if (v.mode === 'multi') {
      const stationData = v.stationData || {};
      for (const stId of Object.keys(stationData)) {
        const s = stationData[stId];
        if (Array.isArray(s.photos)) {
          for (const p of s.photos) {
            if (p && !p.dataUrl && p.id) {
              try { p.dataUrl = await getPhotoDataUrl(p.id); } catch {}
            }
          }
        }
      }
    }
  }
}

// V3からsharedStationsへ移行（既存）
export function migrateSharedStationsFromLegacy(projectDate) {
  const nameSet = new Map();
  Object.keys(dataMap).forEach(cat => {
    const v = dataMap[cat];
    if (v && v.mode === "multi" && Array.isArray(v.stations)) {
      v.stations.forEach(st => {
        const nm = (st.name || "基地局").trim();
        const key = nm.toLowerCase();
        if (!nameSet.has(key)) nameSet.set(key, nm);
      });
    }
  });
  const stations = Array.from(nameSet.values())
    .sort((a,b)=>a.localeCompare(b,"ja"))
    .map(nm => ({ id: stationIdFromName(nm), name: nm }));
  setSharedStations(stations);

  Object.keys(dataMap).forEach(cat => {
    const v = dataMap[cat];
    if (v && v.mode === "multi") {
      const stationData = {};
      if (Array.isArray(v.stations)) {
        v.stations.forEach(st => {
          const nm = (st.name || "基地局").trim();
          const id = stationIdFromName(nm);
          stationData[id] = {
            date: st.date || projectDate || "",
            location: st.location || "",
            details: st.details || "",
            photos: Array.isArray(st.photos) ? st.photos : [],
          };
        });
      }
      stations.forEach(ss => {
        if (!stationData[ss.id]) {
          const tmpl = multiCategoryDefaults[cat] || {};
          stationData[ss.id] = { date: projectDate || "", location: tmpl.location || "", details: tmpl.details || "", photos: [] };
        }
      });
      v.stationData = stationData;
      delete v.stations;
    }
  });
}
