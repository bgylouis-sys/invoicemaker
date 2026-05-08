/* ── State ───────────────────────────────────────────────── */
let itemIndex = 0;
let itemCount = 0;
let activeItemCard = null;
let _productCache = [];

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

function addLogoOptionCard(filename, displayName) {
  const name = displayName || filename.replace(/\.[^.]+$/, '').toUpperCase();
  const container = document.getElementById('logo-presets');
  const label = document.createElement('label');
  label.className = 'logo-option';
  label.dataset.filename = filename;
  label.innerHTML = `<img src="/logo/${filename}" alt="${name}"><span>${name}</span>`;
  label.addEventListener('click', () => selectLogoOption(label));
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
  document.getElementById('pi-form').addEventListener('submit', e => {
    e.preventDefault();
    if (!validateForm()) return;
    submitForm();
  });
}

/* ── Quotation Validity hint sync ────────────────────────── */
function bindValiditySync() {
  const inp  = document.getElementById('quotation_validity');
  const hint = document.getElementById('validity-hint');
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

  // Toggle field visibility on ALL existing item cards
  document.querySelectorAll('.item-card').forEach(card => {
    toggleItemFields(card, isLed);
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
  // Set global project selector
  const globalSel = document.getElementById("global-project-select");
  if (globalSel && productData.project) {
    globalSel.value = productData.project;
  }
  toggleItemFields(card, isLed);
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
  [qtyEl, priceEl].forEach(el => el.addEventListener('input', () => {
    updateCardAmount(card);
    updateSummary();
  }));

  document.getElementById('items-container').appendChild(tpl);

  // Pre-fill from product library if data provided
  if (productData) {
    const isLed = productData.project === 'LED灯具';
    toggleItemFields(card, isLed);
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

function toggleItemFields(card, isLed) {
  const stdFields = card.querySelector('.item-std-fields');
  const ledFields = card.querySelector('.item-led-fields');
  if (stdFields) stdFields.style.display = isLed ? 'none' : '';
  if (ledFields) ledFields.style.display = isLed ? '' : 'none';
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
  const qty   = parseFloat(card.querySelector('[data-fname="item_qty"]').value)   || 0;
  const price = parseFloat(card.querySelector('[data-fname="item_price"]').value) || 0;
  const amt   = qty * price;
  card.querySelector('.amt-val').textContent = formatNum(amt);
  card.dataset.amount = amt;
}

function updateSummary() {
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
  const cur = document.getElementById('currency').value || 'JOD';
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

  const formData = new FormData(document.getElementById('pi-form'));
  try {
    const res = await fetch('/generate', { method: 'POST', body: formData });
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

  const formData = new FormData(document.getElementById('pi-form'));
  try {
    const res = await fetch('/preview', { method: 'POST', body: formData });
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

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('freight')?.addEventListener('input', updateSummary);
  document.getElementById('currency')?.addEventListener('change', updateCurrencyLabel);
});
