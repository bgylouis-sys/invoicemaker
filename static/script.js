/* ── State ───────────────────────────────────────────────── */
let itemIndex = 0;
let itemCount = 0;
let activeItemCard = null;
let _productCache = [];

const APP_MODE = window.APP_MODE || 'pi';
const CFG = APP_MODE === 'ci'
  ? { formId: 'ci-form', generateUrl: '/ci/generate', previewUrl: '/ci/preview' }
  : APP_MODE === 'pl'
    ? { formId: 'pl-form', generateUrl: '/pl/generate', previewUrl: '/pl/preview' }
    : { formId: 'pi-form', generateUrl: '/generate',      previewUrl: '/preview' };

/* ── Init ────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  addItem();
  bindForm();
  bindValiditySync();
  initLogoSelector();
  initPaymentTranslation();
  initProductLibrary();
  initBankInfo();
  initGlobalProject();
  initLEDRemarksPreview();
});

/* ══ Logo selector ═══════════════════════════════════════════ */

function initLogoSelector() {
  const options = document.querySelectorAll('.logo-option');
  options.forEach(opt => opt.addEventListener('click', () => selectLogoOption(opt)));
  if (options.length > 0) selectLogoOption(options[0]);

  document.getElementById('new_logo_input').addEventListener('change', e => {
    const file = e.target.files[0];
    document.getElementById('new-logo-fname').textContent = file ? file.name : '未选择';
  });

  document.getElementById('btn-upload-logo').addEventListener('click', uploadNewLogo);
}

function selectLogoOption(el) {
  document.querySelectorAll('.logo-option').forEach(o => o.classList.remove('selected'));
  el.classList.add('selected');
  document.getElementById('selected_logo_name').value = el.dataset.filename;
}

async function uploadNewLogo() {
  const input     = document.getElementById('new_logo_input');
  const nameInput = document.getElementById('new_logo_name');
  const file = input.files[0];
  if (!file) { showToast('请先选择图片文件 / Please select an image file'); return; }

  const fd = new FormData();
  fd.append('logo_file', file);
  fd.append('brand_name', nameInput.value.trim());

  try {
    const res  = await fetch('/upload-logo', { method: 'POST', body: fd });
    const data = await res.json();
    if (data.filename) {
      addLogoOptionCard(data.filename, data.display_name);
      showToast('Logo 已保存 / Logo saved', true);
      input.value = '';
      nameInput.value = '';
      document.getElementById('new-logo-fname').textContent = '未选择';
    } else {
      showToast('上传失败 / Upload failed: ' + (data.error || ''));
    }
  } catch {
    showToast('上传失败 / Upload failed');
  }
}

async function deleteLogo(btn, filename) {
  if (!confirm(`删除此标志？Delete logo "${filename}"?`)) return;

  try {
    const res = await fetch('/delete-logo', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ filename })
    });
    const data = await res.json();
    if (data.success) {
      const card = btn.closest('.logo-option');
      const wasSelected = card.classList.contains('selected');
      card.remove();
      if (wasSelected) {
        const first = document.querySelector('.logo-option');
        if (first) selectLogoOption(first);
        else document.getElementById('selected_logo_name').value = '';
      }
      showToast('已删除 / Deleted', true);
    } else {
      showToast('删除失败 / Delete failed: ' + (data.error || ''));
    }
  } catch {
    showToast('删除失败 / Delete failed');
  }
}

function addLogoOptionCard(filename, displayName) {
  const name = displayName || filename.replace(/\.[^.]+$/, '').toUpperCase();
  const container = document.getElementById('logo-presets');
  const label = document.createElement('label');
  label.className = 'logo-option';
  label.dataset.filename = filename;
  label.innerHTML =
    `<button type="button" class="logo-del-btn" title="删除 Remove"` +
    ` onclick="event.stopPropagation();deleteLogo(this, '${filename}')">&times;</button>` +
    `<img src="/logo/${filename}" alt="${name}"><span>${name}</span>`;
  label.addEventListener('click', (e) => {
    if (e.target.classList.contains('logo-del-btn')) return;
    selectLogoOption(label);
  });
  container.appendChild(label);
  selectLogoOption(label);
}

function escHtml(str) {
  return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

/* ══ Payment Terms Translation ═══════════════════════════════ */

function initPaymentTranslation() {
  const input = document.getElementById('payment_terms');
  const btn   = document.getElementById('btn-translate-payment');
  const hint  = document.getElementById('translate-hint');
  if (!input) return;

  input.addEventListener('input', () => {
    const hasChinese = /[一-鿿]/.test(input.value);
    btn.style.display = hasChinese ? 'inline-flex' : 'none';
    if (!hasChinese) hint.textContent = '';
  });
}

async function translatePaymentTerms() {
  const input = document.getElementById('payment_terms');
  const btn   = document.getElementById('btn-translate-payment');
  const hint  = document.getElementById('translate-hint');
  const text  = input.value.trim();
  if (!text) return;

  btn.disabled = true;
  btn.textContent = '翻译中…';
  hint.textContent = '';

  try {
    const res  = await fetch('/translate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text }),
    });
    const data = await res.json();
    if (data.translated) {
      input.value = data.translated;
      hint.textContent = '✓ 已翻译为英文 / Translated to English';
      btn.style.display = 'none';
    } else {
      hint.textContent = '翻译失败 / Translation failed: ' + (data.error || '');
    }
  } catch {
    hint.textContent = '翻译服务暂不可用 / Translation service unavailable';
  }

  btn.disabled = false;
  btn.textContent = '🌐 翻译';
}

/* ══ Form submit ═════════════════════════════════════════════ */

function bindForm() {
  document.getElementById(CFG.formId).addEventListener('submit', e => {
    e.preventDefault();
    if (!validateForm()) return;
    submitForm();
  });
}

/* ── Quotation Validity hint sync ────────────────────────── */
function bindValiditySync() {
  const inp  = document.getElementById('quotation_validity');
  const hint = document.getElementById('validity-hint');
  if (!inp || !hint) return;
  inp.addEventListener('input', () => {
    const v = inp.value.trim();
    hint.textContent = v
      ? `将同步填入备注第4条：valid for ${v} days`
      : '输入天数后自动同步到备注第4条';
  });
}

/* ══ Bank Info ════════════════════════════════════════════════ */

let _allBanks = [];

async function initBankInfo() {
  try {
    const res = await fetch('/bank-info');
    _allBanks = await res.json();
  } catch { _allBanks = []; }
  populateBankDropdown();
  document.getElementById('bank-project-select').addEventListener('change', onBankProjectChange);
}

function populateBankDropdown() {
  const sel = document.getElementById('bank-project-select');
  const cur = sel.value;
  sel.innerHTML = '<option value="">-- 请选择项目 / Select Project --</option>';
  _allBanks.forEach(b => {
    const opt = document.createElement('option');
    opt.value = b.project;
    opt.textContent = `${b.project} / ${b.project_en || ''}`;
    sel.appendChild(opt);
  });
  if (cur) sel.value = cur;
}

function onBankProjectChange() {
  const project = document.getElementById('bank-project-select').value;
  const preview = document.getElementById('bank-preview');
  const editForm = document.getElementById('bank-edit-form');
  if (!project) {
    preview.style.display = 'none';
    editForm.style.display = 'none';
    return;
  }
  const bank = _allBanks.find(b => b.project === project);
  if (bank) {
    preview.style.display = 'block';
    preview.querySelectorAll('[data-key]').forEach(el => {
      el.textContent = bank[el.dataset.key] || '';
    });
    document.getElementById('bank-edit-form').style.display = 'none';
  } else {
    preview.style.display = 'none';
    editBankInfo();
  }
}

function editBankInfo() {
  const project = document.getElementById('bank-project-select').value;
  const bank = project ? _allBanks.find(b => b.project === project) : null;
  document.getElementById('bank-edit-project').value = project || '';
  document.getElementById('bank-edit-project-en').value = bank ? (bank.project_en || '') : '';
  document.getElementById('bank-edit-beneficiary').value = bank ? (bank.beneficiary || '') : '';
  document.getElementById('bank-edit-bank').value = bank ? (bank.bank || '') : '';
  document.getElementById('bank-edit-swift').value = bank ? (bank.swift || '') : '';
  document.getElementById('bank-edit-account').value = bank ? (bank.account || '') : '';
  document.getElementById('bank-edit-iban').value = bank ? (bank.iban || '') : '';
  document.getElementById('bank-edit-address').value = bank ? (bank.bank_address || '') : '';
  document.getElementById('bank-edit-project').readOnly = !!bank;
  document.getElementById('bank-preview').style.display = 'none';
  document.getElementById('bank-edit-form').style.display = 'block';
  const delBtn = document.querySelector('.btn-del-bank');
  if (delBtn) delBtn.style.display = (bank && bank.id && !bank.id.startsWith('sanitary') && !bank.id.startsWith('led')) ? 'inline-block' : 'none';
}

function cancelEditBank() {
  document.getElementById('bank-edit-form').style.display = 'none';
  const project = document.getElementById('bank-project-select').value;
  if (project && _allBanks.find(b => b.project === project)) {
    document.getElementById('bank-preview').style.display = 'block';
  }
}

async function saveBankInfo() {
  const project = document.getElementById('bank-edit-project').value.trim();
  if (!project) { showToast('请输入项目名称 / Enter project name'); return; }
  const entry = {
    project:      project,
    project_en:   document.getElementById('bank-edit-project-en').value.trim(),
    beneficiary:  document.getElementById('bank-edit-beneficiary').value.trim(),
    bank:         document.getElementById('bank-edit-bank').value.trim(),
    swift:        document.getElementById('bank-edit-swift').value.trim(),
    account:      document.getElementById('bank-edit-account').value.trim(),
    iban:         document.getElementById('bank-edit-iban').value.trim(),
    bank_address: document.getElementById('bank-edit-address').value.trim(),
  };
  try {
    const res  = await fetch('/save-bank-info', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(entry),
    });
    const data = await res.json();
    if (data.id) {
      _allBanks = _allBanks.filter(b => b.project !== project);
      _allBanks.push(data);
      populateBankDropdown();
      document.getElementById('bank-project-select').value = project;
      document.getElementById('bank-edit-form').style.display = 'none';
      document.getElementById('bank-preview').style.display = 'block';
      onBankProjectChange();
      showToast('银行信息已保存 / Bank info saved', true);
    } else {
      showToast('保存失败 / Save failed: ' + (data.error || ''));
    }
  } catch { showToast('保存失败 / Save failed'); }
}

async function deleteBankInfo() {
  const project = document.getElementById('bank-project-select').value;
  const bank = _allBanks.find(b => b.project === project);
  if (!bank) return;
  try {
    await fetch(`/delete-bank-info/${bank.id}`, { method: 'DELETE' });
    _allBanks = _allBanks.filter(b => b.id !== bank.id);
    populateBankDropdown();
    document.getElementById('bank-project-select').value = '';
    document.getElementById('bank-preview').style.display = 'none';
    document.getElementById('bank-edit-form').style.display = 'none';
    showToast('已删除 / Deleted', true);
  } catch { showToast('删除失败 / Delete failed'); }
}

/* ══ Global Project Selector ══════════════════════════════════ */

function initGlobalProject() {
  const sel = document.getElementById('global-project-select');
  if (!sel) return;
  sel.addEventListener('change', onGlobalProjectChange);
}

function onGlobalProjectChange() {
  const sel = document.getElementById('global-project-select');
  const project = sel ? sel.value : '';
  const isLed = project === 'LED灯具';
  const isTile = project === '瓷砖';

  // Toggle field visibility on ALL existing item cards
  document.querySelectorAll('.item-card').forEach(card => {
    toggleItemFields(card, isLed, isTile);
  });

  // Refresh product library filtered by project
  if (project) {
    refreshProductLibrary(project);
  } else {
    refreshProductLibrary();
  }

  // Toggle LED remarks section
  const ledRemarksCard = document.getElementById('led-remarks-card');
  if (ledRemarksCard) {
    ledRemarksCard.style.display = isLed ? '' : 'none';
  }
}

function getGlobalProject() {
  const sel = document.getElementById('global-project-select');
  return sel ? sel.value : '';
}

/* ══ LED Remarks Live Preview ═══════════════════════════════════ */

function initLEDRemarksPreview() {
  const monthsEl    = document.getElementById('led_delivery_months');
  const depositEl   = document.getElementById('led_deposit_pct');
  const remainingEl = document.getElementById('led_remaining_pct');
  if (!monthsEl) return;

  function updatePreview() {
    const previewMonths    = document.getElementById('led-preview-months');
    const previewDeposit   = document.getElementById('led-preview-deposit');
    const previewRemaining = document.getElementById('led-preview-remaining');
    if (previewMonths)    previewMonths.textContent    = monthsEl.value || '_';
    if (previewDeposit)   previewDeposit.textContent   = depositEl.value || '_';
    if (previewRemaining) previewRemaining.textContent = remainingEl.value || '_';
  }

  monthsEl.addEventListener('input', updatePreview);
  depositEl.addEventListener('input', updatePreview);
  remainingEl.addEventListener('input', updatePreview);
}

/* ══ Product Library ═════════════════════════════════════════ */

const PROJECTS = [
  '卫浴', 'LED灯具', '不锈钢', '床垫', '纸箱', '瓷砖胶',
  '浴室柜', '瓷砖', '注塑', '钢铁厂', '造纸厂',
];

async function initProductLibrary() {
  refreshProductLibrary();
}

async function refreshProductLibrary(project) {
  const filterProject = project || '';
  let url = '/products';
  if (filterProject) url += '?project=' + encodeURIComponent(filterProject);
  try {
    const res  = await fetch(url);
    const list = await res.json();
    _productCache = list;
    renderProductCards(list);
  } catch { /* no products yet */ }
}

function renderProductCards(products) {
  const container = document.getElementById('product-cards');
  const countEl   = document.getElementById('product-count');
  countEl.textContent = products.length ? `${products.length} 个产品` : '';
  if (!products.length) {
    container.innerHTML = '<div class="no-products-msg">暂无保存的产品 / No saved products</div>';
    return;
  }
  container.innerHTML = '';
  products.forEach(p => {
    const card = document.createElement('div');
    card.className = 'product-card';
    const meta = [p.model, p.color || p.dia_shape, p.unit, p.price ? p.price + (p.tax === 'incl' ? '(含税)' : '') : '']
      .filter(Boolean).join(' · ');
    const projectBadge = p.project ? `<span class="prod-project">${escHtml(p.project)}</span>` : '';
    card.innerHTML = `
      <div class="prod-main" data-product-id="${escHtml(p.id)}">
        <div class="prod-name">${escHtml(p.desc)} ${projectBadge}</div>
        <div class="prod-meta">${escHtml(meta)}</div>
      </div>
      <button class="btn-del-product" title="删除" data-del-id="${escHtml(p.id)}">✕</button>`;
    card.querySelector('.prod-main').addEventListener('click', () => fillActiveItem(p));
    card.querySelector('.btn-del-product').addEventListener('click', e => {
      e.stopPropagation();
      deleteProduct(p.id);
    });
    container.appendChild(card);
  });
}

function fillActiveItem(productData) {
  if (!activeItemCard) {
    addItem(productData);
    return;
  }
  const card = activeItemCard;
  const isLed = productData.project === 'LED灯具';
  const isTile = productData.project === '瓷砖';
  // Set global project selector
  const globalSel = document.getElementById("global-project-select");
  if (globalSel && productData.project) {
    globalSel.value = productData.project;
  }
  toggleItemFields(card, isLed, isTile);
  const fieldMap = {
    item_desc:    productData.desc,
    item_model:   productData.model,
    item_unit:    productData.unit,
    item_price:   productData.price,
    item_tax:     productData.tax,
    item_remarks: productData.remarks,
  };
  if (isLed) {
    Object.assign(fieldMap, {
      item_dia_shape:  productData.dia_shape,
      item_lumin:      productData.lumin,
      item_voltage:    productData.voltage,
      item_power:      productData.power,
      item_color_temp: productData.color_temp,
      item_material:   productData.material,
    });
  } else if (isTile) {
    Object.assign(fieldMap, {
      item_size:        productData.size,
      item_tile_type:   productData.tile_type,
      item_thickness:   productData.thickness,
      item_brand:       productData.brand,
      item_m2_per_ctn:  productData.m2_per_ctn,
      item_gw_per_ctn:  productData.gw_per_ctn,
      item_nw_per_ctn:  productData.nw_per_ctn,
    });
  } else {
    Object.assign(fieldMap, {
      item_color: productData.color,
      item_size:  productData.size,
    });
  }
  Object.entries(fieldMap).forEach(([fname, val]) => {
    if (val != null && val !== '') {
      const el = card.querySelector(`[data-fname="${fname}"]`);
      if (el) el.value = val;
    }
  });
  updateCardAmount(card);
  updateSummary();
  refreshProductLibrary(productData.project || "");
}

async function saveProductFromCard(card) {
  const project = getGlobalProject();
  if (!project) { showToast('请先在商品明细中选择项目 / Select a project first'); return; }
  const get = fname => card.querySelector(`[data-fname="${fname}"]`)?.value.trim() || '';
  const product = {
    project: project,
    desc:    get('item_desc'),
    model:   get('item_model'),
    unit:    get('item_unit'),
    price:   get('item_price'),
    tax:     get('item_tax'),
    remarks: get('item_remarks'),
  };
  if (!product.desc) { showToast('请先填写产品描述 / Enter product description first'); return; }
  if (project === 'LED灯具') {
    product.dia_shape  = get('item_dia_shape');
    product.lumin      = get('item_lumin');
    product.voltage    = get('item_voltage');
    product.power      = get('item_power');
    product.color_temp = get('item_color_temp');
    product.material   = get('item_material');
  } else if (project === '瓷砖') {
    product.size        = get('item_size');
    product.tile_type   = get('item_tile_type');
    product.thickness   = get('item_thickness');
    product.brand       = get('item_brand');
    product.m2_per_ctn  = get('item_m2_per_ctn');
    product.gw_per_ctn  = get('item_gw_per_ctn');
    product.nw_per_ctn  = get('item_nw_per_ctn');
  } else {
    product.color = get('item_color');
    product.size  = get('item_size');
  }
  try {
    const res  = await fetch('/save-product', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(product),
    });
    const data = await res.json();
    if (data.id) {
      refreshProductLibrary(project);
      showToast(`已加入产品库：${product.desc}`, true);
    } else {
      showToast('保存失败 / Save failed: ' + (data.error || ''));
    }
  } catch { showToast('保存失败 / Save failed'); }
}

async function deleteProduct(id) {
  try {
    await fetch(`/delete-product/${id}`, { method: 'DELETE' });
    refreshProductLibrary(getGlobalProject());
  } catch { showToast('删除失败 / Delete failed'); }
}

/* ══ Item list ═══════════════════════════════════════════════ */

function addItem(productData = null) {
  const idx = itemIndex++;
  itemCount++;
  document.getElementById('item_count').value = itemCount;

  const tpl  = document.getElementById('item-template').content.cloneNode(true);
  const card = tpl.querySelector('.item-card');
  card.dataset.idx = idx;

  card.querySelectorAll('[data-fname]').forEach(el => {
    const base = el.dataset.fname;
    el.name = `${base}_${idx}`;
    el.id   = `${base}_${idx}`;
  });

  card.querySelectorAll('label[data-for]').forEach(lbl => {
    lbl.setAttribute('for', `${lbl.dataset.for}_${idx}`);
  });

  card.querySelector('.item-num').textContent = `Item #${itemCount}`;
  card.querySelector('.btn-remove').addEventListener('click', () => removeItem(card));
  card.querySelector('.btn-save-product').addEventListener('click', () => saveProductFromCard(card));

  // Click card body → set as active
  card.addEventListener('click', e => {
    if (e.target.closest('button') || e.target.closest('select')) return;
    setActiveItem(card);
  });

  const qtyEl   = card.querySelector('[data-fname="item_qty"]');
  const priceEl = card.querySelector('[data-fname="item_price"]');
  if (APP_MODE === 'pl') {
    const m2El   = card.querySelector('[data-fname="item_m2_per_ctn"]');
    const gwEl   = card.querySelector('[data-fname="item_gw_per_ctn"]');
    const nwEl   = card.querySelector('[data-fname="item_nw_per_ctn"]');
    const pcsEl  = card.querySelector('[data-fname="item_pcs_per_ctn"]');
    const sizeEl = card.querySelector('[data-fname="item_carton_size"]');
    [qtyEl, m2El, gwEl, nwEl, pcsEl, sizeEl].forEach(el => {
      if (el) el.addEventListener('input', () => {
        updateCardAmount(card);
        updateSummary();
      });
    });
  } else {
    [qtyEl, priceEl].forEach(el => {
      if (el) el.addEventListener('input', () => {
        updateCardAmount(card);
        updateSummary();
      });
    });
  }

  document.getElementById('items-container').appendChild(tpl);

  // Pre-fill from product library if data provided
  if (productData) {
    const isLed = productData.project === 'LED灯具';
    const isTile = productData.project === '瓷砖';
    toggleItemFields(card, isLed, isTile);
    const fieldMap = {
      item_desc:    productData.desc,
      item_model:   productData.model,
      item_unit:    productData.unit,
      item_price:   productData.price,
      item_tax:     productData.tax,
      item_remarks: productData.remarks,
    };
    if (isLed) {
      Object.assign(fieldMap, {
        item_dia_shape:  productData.dia_shape,
        item_lumin:      productData.lumin,
        item_voltage:    productData.voltage,
        item_power:      productData.power,
        item_color_temp: productData.color_temp,
        item_material:   productData.material,
      });
    } else if (isTile) {
      Object.assign(fieldMap, {
        item_size:        productData.size,
        item_tile_type:   productData.tile_type,
        item_thickness:   productData.thickness,
        item_brand:       productData.brand,
        item_m2_per_ctn:  productData.m2_per_ctn,
        item_gw_per_ctn:  productData.gw_per_ctn,
        item_nw_per_ctn:  productData.nw_per_ctn,
      });
    } else {
      Object.assign(fieldMap, {
        item_color: productData.color,
        item_size:  productData.size,
      });
    }
    Object.entries(fieldMap).forEach(([fname, val]) => {
      if (val != null && val !== '') {
        const el = card.querySelector(`[data-fname="${fname}"]`);
        if (el) el.value = val;
      }
    });
    updateCardAmount(card);
  }

  setActiveItem(card);
  updateSummary();
}

function setActiveItem(card) {
  if (activeItemCard && activeItemCard !== card) {
    activeItemCard.style.outline = '';
  }
  activeItemCard = card;
  card.style.outline = '2px solid var(--blue)';
  card.style.outlineOffset = '-2px';
  card.style.borderRadius = 'var(--radius)';
}

function toggleItemFields(card, isLed, isTile) {
  const stdFields = card.querySelector('.item-std-fields');
  const ledFields = card.querySelector('.item-led-fields');
  const tileFields = card.querySelector('.item-tile-fields');
  const saniFields = card.querySelector('.item-sanitary-fields');
  if (stdFields) stdFields.style.display = (isLed || isTile) ? 'none' : '';
  if (ledFields) ledFields.style.display = isLed ? '' : 'none';
  if (tileFields) tileFields.style.display = isTile ? '' : 'none';
  if (saniFields) saniFields.style.display = (!isLed && !isTile) ? '' : 'none';
}

function removeItem(card) {
  if (itemCount <= 1) return;
  if (activeItemCard === card) {
    activeItemCard = null;
    card.style.outline = '';
  }
  card.remove();
  itemCount--;
  document.getElementById('item_count').value = itemCount;
  renumberItems();
  updateSummary();
}

function renumberItems() {
  document.querySelectorAll('.item-card').forEach((card, i) => {
    card.querySelector('.item-num').textContent = `Item #${i + 1}`;
  });
}

function updateCardAmount(card) {
  if (APP_MODE === 'pl') {
    const qty = parseFloat(card.querySelector('[data-fname="item_qty"]')?.value) || 0;
    const gw  = parseFloat(card.querySelector('[data-fname="item_gw_per_ctn"]')?.value) || 0;
    const nw  = parseFloat(card.querySelector('[data-fname="item_nw_per_ctn"]')?.value) || 0;
    const project = getGlobalProject();
    const isTile = project === '瓷砖';
    let ctns = 0;
    let cbm = 0;

    if (isTile) {
      const m2 = parseFloat(card.querySelector('[data-fname="item_m2_per_ctn"]')?.value) || 1;
      ctns = m2 > 0 ? qty / m2 : 0;
    } else {
      const pcsPerCtn = parseFloat(card.querySelector('[data-fname="item_pcs_per_ctn"]')?.value) || 0;
      ctns = pcsPerCtn > 0 ? qty / pcsPerCtn : 0;
      // Compute CBM from carton size
      const sizeStr = card.querySelector('[data-fname="item_carton_size"]')?.value || '';
      const dims = sizeStr.match(/(\d+\.?\d*)\s*[×xX]\s*(\d+\.?\d*)\s*[×xX]\s*(\d+\.?\d*)/);
      if (dims) {
        const volCm3 = parseFloat(dims[1]) * parseFloat(dims[2]) * parseFloat(dims[3]);
        cbm = ctns * volCm3 / 1000000;
      }
      const cbmEl = card.querySelector('.cbm-display');
      if (cbmEl) cbmEl.textContent = cbm.toFixed(3);
    }

    card.dataset.ctns     = ctns;
    card.dataset.total_gw = ctns * gw;
    card.dataset.total_nw = ctns * nw;
    card.dataset.cbm      = cbm;
    const ctnsEl = card.querySelector('.amt-val');
    if (ctnsEl) ctnsEl.textContent = Math.round(ctns);
    return;
  }
  const qty   = parseFloat(card.querySelector('[data-fname="item_qty"]').value)   || 0;
  const price = parseFloat(card.querySelector('[data-fname="item_price"]').value) || 0;
  const amt   = qty * price;
  card.querySelector('.amt-val').textContent = formatNum(amt);
  card.dataset.amount = amt;
}

function updateSummary() {
  if (APP_MODE === 'pl') {
    updatePLSummary();
    return;
  }
  let subtotal = 0;
  document.querySelectorAll('.item-card').forEach(card => {
    subtotal += parseFloat(card.dataset.amount || 0);
  });
  const freight = parseFloat(document.getElementById('freight').value) || 0;
  const grand   = subtotal + freight;
  document.getElementById('display-subtotal').textContent = formatNum(subtotal);
  document.getElementById('display-grand').textContent    = formatNum(grand);
  updateCurrencyLabel();
}

function updateCurrencyLabel() {
  const el = document.getElementById('currency');
  const cur = (el && el.value) || 'JOD';
  document.querySelectorAll('.currency-label').forEach(el => el.textContent = cur);
}

/* ── Validation ──────────────────────────────────────────── */
function validateForm() {
  let ok = true;
  const required = [
    { id: 'invoice_no',    label: '发票号 Invoice No.' },
    { id: 'customer_name', label: '客户名称 Customer Name' },
  ];
  required.forEach(({ id, label }) => {
    const el = document.getElementById(id);
    if (!el.value.trim()) {
      el.classList.add('error');
      ok = false;
      el.addEventListener('input', () => el.classList.remove('error'), { once: true });
    }
  });

  const firstDesc = document.querySelector('[data-fname="item_desc"]');
  if (!firstDesc || !firstDesc.value.trim()) {
    showToast('至少填写一行商品描述 / At least one product description required');
    ok = false;
  }

  if (!ok && !firstDesc?.value.trim()) return false;
  if (!ok) showToast('请填写必填字段（红色高亮）/ Please fill required fields');
  return ok;
}

/* ── Submit ──────────────────────────────────────────────── */
async function submitForm(btnId = 'btn-generate') {
  const btn = document.getElementById(btnId);
  btn.disabled = true;
  btn.classList.add('loading');

  const formData = new FormData(document.getElementById(CFG.formId));
  try {
    const res = await fetch(CFG.generateUrl, { method: 'POST', body: formData });
    if (!res.ok) {
      let msg = '生成失败 / Generation failed';
      try { const j = await res.json(); msg = j.error || msg; } catch {}
      throw new Error(msg);
    }
    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    const cd   = res.headers.get('Content-Disposition') || '';
    const m    = cd.match(/filename\*?=(?:UTF-8'')?["']?([^"';\n]+)/i);
    a.download  = m ? decodeURIComponent(m[1].trim()) : 'invoice.xlsx';
    a.href = url;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast('✓ Excel 已生成下载 / Excel downloaded', true);
  } catch (e) {
    showToast('生成失败 / Failed: ' + e.message);
  } finally {
    btn.disabled = false;
    btn.classList.remove('loading');
  }
}

/* ── Export PDF ──────────────────────────────────────────── */
async function exportPDF(btnId = 'btn-export-pdf') {
  const btn = document.getElementById(btnId);
  btn.disabled = true;
  const orig = btn.textContent;
  btn.textContent = '处理中… / Processing…';

  const formData = new FormData(document.getElementById(CFG.formId));
  try {
    const res = await fetch(CFG.previewUrl, { method: 'POST', body: formData });
    if (!res.ok) throw new Error('Preview failed');
    const html = await res.text();
    const win = window.open('', '_blank');
    win.document.write(html);
    win.document.close();
  } catch (e) {
    showToast('PDF 导出失败 / PDF export failed: ' + e.message);
  } finally {
    btn.disabled = false;
    btn.textContent = orig;
  }
}

/* ── Arabic (EN+AR) export wrappers ──────────────────────── */
async function generateAR() {
  document.getElementById('form-lang').value = 'ar';
  await submitForm('btn-generate-ar');
  document.getElementById('form-lang').value = 'zh';
}

async function exportPDFAR() {
  document.getElementById('form-lang').value = 'ar';
  await exportPDF('btn-export-pdf-ar');
  document.getElementById('form-lang').value = 'zh';
}

/* ── Toast ───────────────────────────────────────────────── */
function showToast(msg, success = false) {
  const existing = document.querySelector('.toast');
  if (existing) existing.remove();
  const t = document.createElement('div');
  t.className = 'toast' + (success ? ' success' : '');
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 4000);
}

/* ── Utility ─────────────────────────────────────────────── */
function formatNum(n) {
  return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/* ══ Transport Mode (PL Maker) ═══════════════════════════════ */

function onTransportModeChange() {
  const mode = document.getElementById('transport-mode');
  if (!mode) return;
  const isSea = mode.value === 'sea';

  // Update port label
  const portLabel = document.getElementById('port-label');
  const portZh = document.getElementById('port-zh');
  if (portLabel) {
    portLabel.innerHTML = isSea
      ? 'Port of Loading &nbsp;<span class="zh" id="port-zh">装运港</span>'
      : 'Loading Place &nbsp;<span class="zh" id="port-zh">装运地</span>';
  }

  // Update CTN label in ALL item cards
  document.querySelectorAll('.item-card').forEach(card => {
    const ctnLabel = card.querySelector('[data-for="item_ctn_no"]');
    if (ctnLabel) {
      ctnLabel.innerHTML = isSea
        ? 'CTN No. &nbsp;<span class="zh">柜号</span>'
        : 'Truck Plate &nbsp;<span class="zh">车牌号</span>';
    }
  });
}

/* ── PL Summary (packing-based) ──────────────────────────── */
function updatePLSummary() {
  let totalCtns = 0;
  let totalGw = 0;
  let totalNw = 0;
  let totalCbm = 0;
  document.querySelectorAll('.item-card').forEach(card => {
    totalCtns += parseFloat(card.dataset.ctns || 0);
    totalGw   += parseFloat(card.dataset.total_gw || 0);
    totalNw   += parseFloat(card.dataset.total_nw || 0);
    totalCbm  += parseFloat(card.dataset.cbm || 0);
  });
  const ctnsEl = document.getElementById('display-total-ctns');
  const gwEl   = document.getElementById('display-total-gw');
  const nwEl   = document.getElementById('display-total-nw');
  const cbmEl  = document.getElementById('display-total-cbm');
  if (ctnsEl) ctnsEl.textContent = Math.round(totalCtns);
  if (gwEl)   gwEl.textContent   = totalGw.toFixed(2);
  if (nwEl)   nwEl.textContent   = totalNw.toFixed(2);
  if (cbmEl)  cbmEl.textContent  = totalCbm.toFixed(3);
}

function updatePLCardAmounts(card) {
  const qty = parseFloat(card.querySelector('[data-fname="item_qty"]')?.value) || 0;
  const m2  = parseFloat(card.querySelector('[data-fname="item_m2_per_ctn"]')?.value) || 1;
  const gw  = parseFloat(card.querySelector('[data-fname="item_gw_per_ctn"]')?.value) || 0;
  const nw  = parseFloat(card.querySelector('[data-fname="item_nw_per_ctn"]')?.value) || 0;
  const ctns = m2 > 0 ? qty / m2 : 0;
  card.dataset.ctns     = ctns;
  card.dataset.total_gw = ctns * gw;
  card.dataset.total_nw = ctns * nw;
}

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('freight')?.addEventListener('input', updateSummary);
  document.getElementById('currency')?.addEventListener('change', updateCurrencyLabel);
  if (APP_MODE === 'pl') {
    onTransportModeChange();
  }
});
