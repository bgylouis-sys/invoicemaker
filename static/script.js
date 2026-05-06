/* ── State ───────────────────────────────────────────────── */
let itemIndex = 0;
let itemCount = 0;

/* ── Init ────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  addItem();
  bindForm();
  bindValiditySync();
  initLogoSelector();
  initPaymentTranslation();
  initProductLibrary();
});

/* ══ Logo selector ═══════════════════════════════════════════ */

function initLogoSelector() {
  const options = document.querySelectorAll('.logo-option');
  options.forEach(opt => opt.addEventListener('click', () => selectLogoOption(opt)));
  if (options.length > 0) selectLogoOption(options[0]);

  // Show selected filename in hint when file chosen
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

/* ══ Product Library ═════════════════════════════════════════ */

const PROJECTS = [
  '卫浴', 'LED灯具', '不锈钢', '床垫', '纸箱', '瓷砖胶',
  '浴室柜', '瓷砖', '注塑', '钢铁厂', '造纸厂',
];

async function initProductLibrary() {
  const filter = document.getElementById('product-project-filter');
  filter.addEventListener('change', () => refreshProductLibrary());
  refreshProductLibrary();
}

async function refreshProductLibrary() {
  const project = document.getElementById('product-project-filter').value;
  let url = '/products';
  if (project) url += '?project=' + encodeURIComponent(project);
  try {
    const res  = await fetch(url);
    const list = await res.json();
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
    const meta = [p.model, p.color, p.unit, p.price ? p.price + (p.tax === 'incl' ? '(含税)' : '') : '']
      .filter(Boolean).join(' · ');
    const projectBadge = p.project ? `<span class="prod-project">${escHtml(p.project)}</span>` : '';
    card.innerHTML = `
      <div class="prod-main" onclick='addItem(${JSON.stringify(p)})'>
        <div class="prod-name">${escHtml(p.desc)} ${projectBadge}</div>
        <div class="prod-meta">${escHtml(meta)}</div>
      </div>
      <button class="btn-del-product" title="删除" onclick="deleteProduct('${p.id}', event)">✕</button>`;
    container.appendChild(card);
  });
}

async function saveProductFromCard(card) {
  const project = document.getElementById('product-project-filter').value;
  if (!project) { showToast('请先在产品库中选择项目 / Select a project in Product Library first'); return; }
  const get = fname => card.querySelector(`[name^="${fname}_"]`)?.value.trim() || '';
  const product = {
    project:   project,
    desc:      get('item_desc'),
    model:     get('item_model'),
    color:     get('item_color'),
    size:      get('item_size'),
    unit:      get('item_unit'),
    price:     get('item_price'),
    tax:       get('item_tax'),
    remarks:   get('item_remarks'),
  };
  if (!product.desc) { showToast('请先填写产品描述 / Enter product description first'); return; }
  try {
    const res  = await fetch('/save-product', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(product),
    });
    const data = await res.json();
    if (data.id) {
      refreshProductLibrary();
      showToast(`已加入产品库：${product.desc}`, true);
    } else {
      showToast('保存失败 / Save failed: ' + (data.error || ''));
    }
  } catch { showToast('保存失败 / Save failed'); }
}

async function deleteProduct(id, event) {
  event.stopPropagation();
  try {
    await fetch(`/delete-product/${id}`, { method: 'DELETE' });
    await initProductLibrary();
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

  const qtyEl   = card.querySelector('[data-fname="item_qty"]');
  const priceEl = card.querySelector('[data-fname="item_price"]');
  [qtyEl, priceEl].forEach(el => el.addEventListener('input', () => {
    updateCardAmount(card);
    updateSummary();
  }));

  document.getElementById('items-container').appendChild(tpl);

  // Pre-fill from product library if data provided (card is now in DOM)
  if (productData) {
    const fieldMap = {
      item_desc:    productData.desc,
      item_model:   productData.model,
      item_color:   productData.color,
      item_size:    productData.size,
      item_unit:    productData.unit,
      item_price:   productData.price,
      item_tax:     productData.tax,
      item_remarks: productData.remarks,
    };
    Object.entries(fieldMap).forEach(([fname, val]) => {
      if (val != null && val !== '') {
        const el = document.getElementById(`${fname}_${idx}`);
        if (el) el.value = val;
      }
    });
    updateCardAmount(card);
  }

  updateSummary();
}

function removeItem(card) {
  if (itemCount <= 1) return;
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
  const qty   = parseFloat(card.querySelector('[name^="item_qty_"]').value)   || 0;
  const price = parseFloat(card.querySelector('[name^="item_price_"]').value) || 0;
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

  const firstDesc = document.querySelector('[name^="item_desc_"]');
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
