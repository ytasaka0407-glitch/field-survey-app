// FieldSurveyApp/assets/js/modules/ui/fields.js
import { readFileAsDataURL, resizeImage } from '../utils.js';

function setNgReasonVisibility(container, model) {
  const el = container.querySelector('[data-field-key="diagramNgReason"]');
  if (el) el.style.display = model.diagramStatus === 'ng' ? '' : 'none';
}
function setMethodVisibility(container, model) {
  const el = container.querySelector('[data-field-key="method"]');
  if (el) el.style.display = (model.installType === 'new') ? '' : 'none';
}

export const FieldRenderers = {
  date: (container, model, field, fid) => {
    container.insertAdjacentHTML('beforeend', `
      <div class="form-row" data-field-key="${field.key}">
        <label for="${fid}">${field.label}</label>
        <input id="${fid}" type="date" value="${model[field.key] || ''}" />
      </div>`);
    container.querySelector(`#${fid}`)
      .addEventListener('change', e => (model[field.key] = e.target.value));
  },

  text: (container, model, field, fid) => {
    container.insertAdjacentHTML('beforeend', `
      <div class="form-row" data-field-key="${field.key}">
        <label for="${fid}">${field.label}</label>
        <input id="${fid}" type="text" placeholder="${field.placeholder || ''}" value="${model[field.key] || ''}" />
      </div>`);
    container.querySelector(`#${fid}`)
      .addEventListener('input', e => (model[field.key] = e.target.value));
  },

  number: (container, model, field, fid) => {
    container.insertAdjacentHTML('beforeend', `
      <div class="form-row" data-field-key="${field.key}">
        <label for="${fid}">${field.label}</label>
        <input id="${fid}" type="number" step="any" value="${model[field.key] ?? ''}" />
      </div>`);
    container.querySelector(`#${fid}`)
      .addEventListener('input', e => (model[field.key] = e.target.value));
  },

  textarea: (container, model, field, fid) => {
    container.insertAdjacentHTML('beforeend', `
      <div class="form-row" data-field-key="${field.key}">
        <label for="${fid}">${field.label}</label>
        <textarea id="${fid}" placeholder="${field.placeholder || ''}">${model[field.key] || ''}</textarea>
      </div>`);
    container.querySelector(`#${fid}`)
      .addEventListener('input', e => (model[field.key] = e.target.value));
  },

  radio: (container, model, field, fidBase) => {
    const opts = Array.isArray(field.options) ? field.options : [];
    const name = `radio_${fidBase}`;
    const value = model[field.key] ?? (opts[0]?.value ?? '');
    if (model[field.key] == null) model[field.key] = value;

    const radios = opts.map((o, i) => {
      const id = `${fidBase}_${i}`;
      const checked = value === o.value ? 'checked' : '';
      return `
        <label style="display:inline-flex; align-items:center; gap:6px; margin-right:12px;">
          <input type="radio" name="${name}" id="${id}" value="${o.value}" ${checked} />
          <span>${o.label}</span>
        </label>`;
    }).join('');

    container.insertAdjacentHTML('beforeend', `
      <div class="form-row" data-field-key="${field.key}">
        <label>${field.label}</label>
        <div>${radios}</div>
      </div>`);

    container.querySelectorAll(`input[name="${name}"]`).forEach(inp => {
      inp.addEventListener('change', (e) => {
        if (e.target.checked) {
          model[field.key] = e.target.value;
          if (field.key === 'diagramStatus') {
            setNgReasonVisibility(container, model);
          } else if (field.key === 'installType') {
            setMethodVisibility(container, model);
          }
        }
      });
    });
  },

  photos: (container, model, field, fidBase) => {
    const key = field.key;
    if (!Array.isArray(model[key])) model[key] = [];

    container.insertAdjacentHTML('beforeend', `
      <div class="form-row" data-field-key="${field.key}">
        <label>${field.label}</label>
        <div>
          <div class="buttons">
            <input id="cam_${fidBase}" type="file" accept="image/*" capture="environment" multiple style="display:none" />
            <button id="camBtn_${fidBase}" class="btn primary">撮影する（カメラ）</button>
            <input id="file_${fidBase}" type="file" accept="image/*" multiple style="display:none" />
            <button id="fileBtn_${fidBase}" class="btn">画像を選択</button>
          </div>
          <div class="photos">
            <div id="list_${fidBase}" class="photo-list"></div>
          </div>
        </div>
      </div>`);

    const camInput  = container.querySelector(`#cam_${fidBase}`);
    const fileInput = container.querySelector(`#file_${fidBase}`);
    container.querySelector(`#camBtn_${fidBase}`).addEventListener('click', () => camInput.click());
    container.querySelector(`#fileBtn_${fidBase}`).addEventListener('click', () => fileInput.click());

    const renderPhotoList = () => {
      const listEl = container.querySelector(`#list_${fidBase}`);
      listEl.innerHTML = '';
      model[key].forEach((p, idx) => {
        const item = document.createElement('div');
        item.className = 'photo-item';
        item.innerHTML = `
          <img src="${p.dataUrl}" alt="photo ${idx + 1}">
          <div class="caption">
            <textarea placeholder="画像の説明を入力">${p.caption || ''}</textarea>
          </div>
          <div class="footer">
            <span>${p.name || ('photo_' + (idx + 1) + '.jpg')}</span>
            <button class="remove">削除</button>
          </div>`;
        const ta = item.querySelector('textarea');
        ta.addEventListener('input', e => { p.caption = e.target.value; });
        item.querySelector('.remove').addEventListener('click', () => {
          model[key].splice(idx, 1);
          renderPhotoList();
        });
        listEl.appendChild(item);
      });
    };

    async function handleFiles(fileList) {
      for (const file of fileList) {
        if (!file.type.startsWith('image/')) continue;
        const dataUrl = await readFileAsDataURL(file);
        const resized = await resizeImage(dataUrl, 1024, 0.8);
        model[key].push({ dataUrl: resized, name: file.name, caption: '' });
      }
      renderPhotoList();
    }

    camInput.addEventListener('change', e => handleFiles(e.target.files));
    fileInput.addEventListener('change', e => handleFiles(e.target.files));

    renderPhotoList();
  },
};

export function renderField(container, model, field, fid) {
  const fn = FieldRenderers[field.type];
  if (!fn) return;
  fn(container, model, field, fid);
  // 初期表示時の表示制御
  if (field.key === 'diagramStatus' || field.key === 'diagramNgReason') {
    setNgReasonVisibility(container, model);
  }
  if (field.key === 'installType' || field.key === 'method') {
    setMethodVisibility(container, model);
  }
}

