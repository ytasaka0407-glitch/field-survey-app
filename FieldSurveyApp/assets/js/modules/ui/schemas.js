// FieldSurveyApp/assets/js/modules/ui/schemas.js
// Single/Multi共通の基本フィールド
export const baseSingleSchema = [
  { key: 'date', type: 'date', label: '調査日', required: true },
  { key: 'location', type: 'text', label: '設置場所', placeholder: '設置する箇所をなるべく詳細に記載' },
  { key: 'details', type: 'textarea', label: '調査内容', placeholder: '現地での調査事項、寸法、注意点など' },
  { key: 'photos', type: 'photos', label: '写真' },
];

export const baseMultiSchema = [
  { key: 'date', type: 'date', label: '調査日', required: true },
  { key: 'location', type: 'text', label: '設置場所', placeholder: '設置する箇所をなるべく詳細に記載' },
  { key: 'details', type: 'textarea', label: '調査内容', placeholder: '現地での調査事項、寸法、注意点など' },
  { key: 'photos', type: 'photos', label: '写真' },
];

// カテゴリ固有の上書き（必要時にここへ追加）
// 例: '07.基地局_アンテナ' など
export const categorySchemaOverrides = {
  // '07.基地局_アンテナ': [
  //   { key: 'date', type: 'date', label: '調査日', required: true },
  //   { key: 'type', type: 'text', label: '型式' },
  //   { key: 'direction', type: 'text', label: '方向' },
  //   { key: 'height', type: 'number', label: '高さ(m)' },
  //   { key: 'mount', type: 'text', label: '支持方法' },
  //   { key: 'condition', type: 'textarea', label: '劣化状況' },
  //   { key: 'photos', type: 'photos', label: '写真' },
  // ],
   '07.基地局_アンテナ': [
     { key: 'date', type: 'date', label: '調査日', required: true },
     { key: 'type', type: 'text', label: '型式' },
     { key: 'direction', type: 'text', label: '方向' },
     { key: 'height', type: 'number', label: '高さ(m)' },
     { key: 'mount', type: 'text', label: '支持方法' },
     { key: 'condition', type: 'textarea', label: '劣化状況' },
     { key: 'photos', type: 'photos', label: '写真' },
   ],
};

// モードとカテゴリ名からスキーマを返す
export function getSchemaFor(cat, mode) {
  const override = categorySchemaOverrides[cat];
  if (override && Array.isArray(override)) return override;
  return mode === 'multi' ? baseMultiSchema : baseSingleSchema;

}
