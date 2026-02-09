import { dataMap, selectedCategories, sharedStations, setSharedStations } from './state.js';
import { stationIdFromName } from './utils.js';
import { multiCategoryDefaults } from './state.js';

export function saveDraft(projectTitle, projectDate) {
  const payload = {
    projectTitle: projectTitle || "",
    projectDate: projectDate || "",
    selected: [...selectedCategories],
    data: dataMap,
    sharedStations,
  };
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

// V3からsharedStationsへ移行
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