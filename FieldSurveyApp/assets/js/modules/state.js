export const defaultSingleCategories = ["01.無線回線制御装置", "02.高機能遠隔制御装置", "03.管理監視制御卓", "04.プリンタ", "05.L3SW(無線系)", "06.無線指令受付装置", "07.署所端末受令機", "08.避雷器", "09.共用器", "10.アンテナ", "11.DC/ACインバータ", "12.非常用発電機", "13.直流電源装置", "14.無停電電源装置"];
export const defaultMultiCategories  = ["01.基地局_無線装置", "02.基地局_L3SW(無線系)", "03.基地局_卓上型固定移動局", "04.基地局_署所端末受令機", "05.基地局_避雷器", "06.基地局_共用器", "07.基地局_アンテナ", "08.基地局_DC/ACインバータ", "09.基地局_非常用発電機", "10.基地局_直流電源装置", "11.基地局_無停電電源装置", "12.基地局_耐雷トランス"];

export const multiCategoryDefaults = {}; // 必要に応じて設定

export const dataMap = {};
export const selectedCategories = new Set();
export let sharedStations = [];
export let projectDatePrev = "";

export function initDefaults() {
  defaultSingleCategories.forEach((c) => {
    if (!dataMap[c]) dataMap[c] = { mode: "single", date: "", location: "", details: "", photos: [] };
  });
  defaultMultiCategories.forEach((c) => {
    if (!dataMap[c]) dataMap[c] = { mode: "multi", stationData: {} };
  });
}

export function ensureSingle(cat, projectDate) {
  const cur = dataMap[cat];
  if (!cur) {
    dataMap[cat] = { mode: "single", date: projectDate || "", location: "", details: "", photos: [] };
  } else if (cur.mode === undefined) {
    cur.mode = "single";
    cur.date = cur.date || projectDate || "";
    cur.location = cur.location || "";
    cur.details = cur.details || "";
    cur.photos = Array.isArray(cur.photos) ? cur.photos : [];
  }
  return dataMap[cat];
}
export function ensureMulti(cat, projectDate) {
  const cur = dataMap[cat];
  if (!cur || cur.mode !== "multi") {
    dataMap[cat] = { mode: "multi", stationData: {} };
  } else if (!cur.stationData) {
    cur.stationData = {};
  }
  sharedStations.forEach(st => {
    if (!dataMap[cat].stationData[st.id]) {
      dataMap[cat].stationData[st.id] = {
        date: projectDate || "",
        location: "",
        details: "",
        photos: [],
      };
    }
  });
  return dataMap[cat];
}
export function getOrInitStationData(cat, stationId, projectDate) {
  const v = ensureMulti(cat, projectDate);
  if (!v.stationData[stationId]) {
    v.stationData[stationId] = {
      date: projectDate || "",
      location: "",
      details: "",
      photos: [],
    };
  }
  return v.stationData[stationId];
}
export function setSharedStations(list) { sharedStations = list; }
export function addSharedStation(st) { sharedStations.push(st); }
export function removeSharedStationById(stationId) {
  const idx = sharedStations.findIndex(s => s.id === stationId);
  if (idx >= 0) sharedStations.splice(idx, 1);
  Object.keys(dataMap).forEach(cat => {
    const v = dataMap[cat];
    if (v?.mode === "multi" && v.stationData?.[stationId]) {
      delete v.stationData[stationId];
    }
  });
}
export function setProjectDatePrev(v) { projectDatePrev = v || ""; }