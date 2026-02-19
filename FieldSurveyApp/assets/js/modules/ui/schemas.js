// FieldSurveyApp/assets/js/modules/ui/schemas.js
// Single/Multi共通の基本フィールド
export const baseSingleSchema = [
  { key: 'date', type: 'date', label: '調査日', required: true },
  { key: 'location', type: 'text', label: '設置場所', placeholder: '設置する箇所をなるべく詳細に記載' },
  { key: 'installType', type: 'radio', label: '新設/既設流用', options: [
    { value: 'new',   label: '新設' },
    { value: 'reuse', label: '既設流用' },
  ]},
  { key: 'method', type: 'text', label: '設置方法', placeholder: '設置方法を記載　例)既設位置に入替など' },
  { key: 'diagramStatus', type: 'radio', label: '系統図との整合性', options: [
    { value: 'ok', label: 'OK' },
    { value: 'ng', label: 'NG' },
  ]},
  { key: 'diagramNgReason', type: 'textarea', label: 'NG理由', placeholder: 'NGの理由を記載' },
  { key: 'details', type: 'textarea', label: 'その他調査内容', placeholder: '現地での調査事項、寸法、注意点など' },
  { key: 'photos', type: 'photos', label: '写真' },
];

export const baseMultiSchema = [
  { key: 'date', type: 'date', label: '調査日', required: true },
  { key: 'location', type: 'text', label: '設置場所', placeholder: '設置する箇所をなるべく詳細に記載' },
  { key: 'installType', type: 'radio', label: '新設/既設流用', options: [
    { value: 'new',   label: '新設' },
    { value: 'reuse', label: '既設流用' },
  ]},
  { key: 'method', type: 'text', label: '設置方法', placeholder: '設置方法を記載　例)既設位置に入替など' },
  { key: 'diagramStatus', type: 'radio', label: '系統図との整合性', options: [
    { value: 'ok', label: 'OK' },
    { value: 'ng', label: 'NG' },
  ]},
  { key: 'diagramNgReason', type: 'textarea', label: 'NG理由', placeholder: 'NGの理由を記載' },
  { key: 'details', type: 'textarea', label: 'その他調査内容', placeholder: '現地での調査事項、寸法、注意点など' },
  { key: 'photos', type: 'photos', label: '写真' },
];

// カテゴリ固有の上書き（必要時にここへ追加）
export const categorySchemaOverrides = {
  // 例:
  // '07.基地局_アンテナ': [ ... ]
};

// モードとカテゴリ名からスキーマを返す
export function getSchemaFor(cat, mode) {
  const override = categorySchemaOverrides[cat];
  if (override && Array.isArray(override)) return override;
  return mode === 'multi' ? baseMultiSchema : baseSingleSchema;
}


