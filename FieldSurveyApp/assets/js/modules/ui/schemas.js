// FieldSurveyApp/assets/js/modules/ui/schemas.js

// Single/Multi共通の基本フィールド（details/photos を分離して差し込みやすくする）
const baseCommonFields = [
  { key: 'date', type: 'date', label: '調査日', required: true },
  { key: 'location', type: 'text', label: '設置場所', placeholder: '設置する箇所をなるべく詳細に記載' },
  {
    key: 'installType',
    type: 'radio',
    label: '新設/既設流用',
    options: [
      { value: 'new', label: '新設' },
      { value: 'reuse', label: '既設流用' },
    ],
  },
  { key: 'method', type: 'text', label: '設置方法', placeholder: '設置方法を記載　例)既設位置に入替など' },
  {
    key: 'diagramStatus',
    type: 'radio',
    label: '系統図との整合性',
    options: [
      { value: 'ok', label: 'OK' },
      { value: 'ng', label: 'NG' },
    ],
  },
  { key: 'diagramNgReason', type: 'textarea', label: 'NG理由', placeholder: 'NGの理由を記載' },
];

// 末尾ブロック（ここに「差し込みたい項目」を挟める）
export const detailsField = {
  key: 'details',
  type: 'textarea',
  label: 'その他調査内容',
  placeholder: '現地での調査事項、寸法、注意点など',
};

export const photosField = { key: 'photos', type: 'photos', label: '写真' };

// Single/Multi のベース（details と photos を別配列として切り離し）
export const baseSingleSchema = [
  ...baseCommonFields,
  detailsField,
  photosField,
];

export const baseMultiSchema = [
  ...baseCommonFields,
  detailsField,
  photosField,
];

// カテゴリ固有の上書き（必要時にここへ追加）
export const categorySchemaOverrides = {
  // 例：details と photos の間に差し込みたい場合
  // '07.基地局_アンテナ': [
  //   ...baseCommonFields,
  //   detailsField,
  //   { key: 'antennaHeight', type: 'text', label: 'アンテナ高', placeholder: 'm' },
  //   photosField,
  // ],
};

// モードとカテゴリ名からスキーマを返す
export function getSchemaFor(cat, mode) {
  const override = categorySchemaOverrides[cat];
  if (override && Array.isArray(override)) return override;
  return mode === 'multi' ? baseMultiSchema : baseSingleSchema;
}
