from flask import Flask, render_template, request, send_file, jsonify, send_from_directory
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from copy import copy
import io, os, json, uuid, urllib.request, urllib.parse
from datetime import date, datetime
from PIL import Image as PILImage

app = Flask(__name__)

BASE_DIR        = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH   = os.path.join(BASE_DIR, '约旦卫浴Proforma Invoice模板.xlsx')
LOGO_DIR        = os.path.join(BASE_DIR, 'logo')
LOGO_NAMES_PATH = os.path.join(LOGO_DIR, 'names.json')
CUSTOMERS_PATH  = os.path.join(BASE_DIR, 'customers.json')
PRODUCTS_PATH   = os.path.join(BASE_DIR, 'products.json')
LOGO_EXTS       = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'}

ITEMS_START = 17
ITEMS_END   = 21
ITEMS_SLOTS = 5


# ── Logo helpers ──────────────────────────────────────────────────────────────

def list_logo_files():
    return sorted(
        f for f in os.listdir(LOGO_DIR)
        if os.path.splitext(f)[1].lower() in LOGO_EXTS
    )

def load_logo_names():
    if os.path.exists(LOGO_NAMES_PATH):
        with open(LOGO_NAMES_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_logo_names(names):
    with open(LOGO_NAMES_PATH, 'w', encoding='utf-8') as f:
        json.dump(names, f, ensure_ascii=False, indent=2)

def logo_display_name(filename, names):
    return names.get(filename, os.path.splitext(filename)[0].upper())

def logo_list_with_names():
    files = list_logo_files()
    names = load_logo_names()
    return [{'filename': f, 'display_name': logo_display_name(f, names)} for f in files]


# ── Customer helpers ──────────────────────────────────────────────────────────

def load_customers():
    if os.path.exists(CUSTOMERS_PATH):
        with open(CUSTOMERS_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def save_customers_file(customers):
    with open(CUSTOMERS_PATH, 'w', encoding='utf-8') as f:
        json.dump(customers, f, ensure_ascii=False, indent=2)


# ── Product helpers ───────────────────────────────────────────────────────────

def load_products():
    if os.path.exists(PRODUCTS_PATH):
        with open(PRODUCTS_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def save_products_file(products):
    with open(PRODUCTS_PATH, 'w', encoding='utf-8') as f:
        json.dump(products, f, ensure_ascii=False, indent=2)


# ── Excel helpers ─────────────────────────────────────────────────────────────

def copy_row_style(ws, src_row, dst_row, max_col=14):
    for c in range(1, max_col + 1):
        src = ws.cell(src_row, c)
        dst = ws.cell(dst_row, c)
        if src.has_style:
            dst.font      = copy(src.font)
            dst.fill      = copy(src.fill)
            dst.border    = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format


def insert_image(ws, file_obj, anchor, w=100, h=75):
    try:
        raw = file_obj.read()
        pil = PILImage.open(io.BytesIO(raw))
        pil.thumbnail((w, h), PILImage.LANCZOS)
        buf = io.BytesIO()
        fmt = (pil.format or 'PNG').upper()
        if fmt not in ('PNG', 'JPEG', 'GIF', 'BMP'):
            fmt = 'PNG'
        if fmt == 'JPEG' and pil.mode == 'RGBA':
            pil = pil.convert('RGB')
        pil.save(buf, format=fmt)
        buf.seek(0)
        xl_img = XLImage(buf)
        xl_img.anchor = anchor
        ws.add_image(xl_img)
    except Exception as e:
        print(f'[image] {anchor}: {e}')


def fill_template(form, files, lang='zh'):
    with open(TEMPLATE_PATH, 'rb') as f:
        buf = io.BytesIO(f.read())
    wb = openpyxl.load_workbook(buf)
    ws = wb.active

    is_ar = (lang == 'ar')
    n = max(1, int(form.get('item_count') or 1))

    # ── Supplier info (rows 4-8 left, value col=C) ───────────────────────────
    ws.cell(4, 3).value = form.get('company_name', '')
    ws.cell(5, 3).value = form.get('address', '')
    ws.cell(6, 3).value = form.get('city_country', '')
    tel = (form.get('tel') or '').strip()
    ws.cell(7, 3).value = tel if tel else None
    ws.cell(8, 3).value = form.get('email', '')

    # ── Invoice info + key terms (rows 4-8 right, value col=J) ──────────────
    ws.cell(4, 10).value = form.get('invoice_no', '')
    date_raw = form.get('date', '')
    if date_raw and '-' in date_raw:
        try:
            d_obj = datetime.strptime(date_raw, '%Y-%m-%d')
            date_str = f"{d_obj.year}.{d_obj.month}.{d_obj.day}"
        except ValueError:
            date_str = date_raw
    else:
        date_str = date_raw
    ws.cell(5, 10).value = date_str
    currency = form.get('currency') or 'JOD'
    ws.cell(6, 10).value = currency
    ws.cell(7, 10).value = form.get('incoterms', '')
    ws.cell(8, 10).value = form.get('payment_terms', '')

    # ── Customer info (rows 10-14 left, value col=C) ──────────────────────────
    ws.cell(10, 3).value = form.get('customer_name', '')
    ws.cell(11, 3).value = form.get('customer_address', '')
    ws.cell(12, 3).value = form.get('customer_email', '')
    ws.cell(13, 3).value = form.get('country', '')
    cust_tel = (form.get('customer_tel') or '').strip()
    ws.cell(14, 3).value = cust_tel if cust_tel else None

    # ── Quotation details (rows 10-14 right, value col=J) ──────────────────────
    prefix = (form.get('contact_prefix') or '').strip()
    person = (form.get('contact_person') or '').strip()
    ws.cell(10, 10).value = f"{prefix} {person}".strip()
    ws.cell(11, 10).value = form.get('inquiry_ref', '')
    ws.cell(12, 10).value = form.get('lead_time', '')
    validity = (form.get('quotation_validity') or '').strip()
    ws.cell(13, 10).value = f'{validity} days' if validity else ''
    port = (form.get('port_of_loading') or '').strip()
    dest = (form.get('destination') or '').strip()
    if port and dest:    shipping = f'{port} → {dest}'
    elif port:          shipping = port
    else:               shipping = dest
    ws.cell(14, 10).value = shipping

    # ── Overwrite section headers for Arabic ──────────────────────────────────
    if is_ar:
        ws.cell(3,  1).value = 'FROM / SUPPLIER  المورّد'
        ws.cell(3,  7).value = 'INVOICE INFO  معلومات الفاتورة'
        ws.cell(9,  1).value = 'TO / CUSTOMER  العميل'
        ws.cell(9,  7).value = 'QUOTATION DETAILS  تفاصيل العرض'
        ws.cell(15, 1).value = 'ITEM LIST  قائمة المنتجات'
        ws.cell(25, 1).value = 'REMARKS  ملاحظات'
        ws.cell(31, 1).value = 'TRANSFER DETAILS  معلومات التحويل'
        ws.cell(36, 1).value = 'Company Stamp  ختم الشركة:'

    # ── Insert extra item rows if n > 5 ─────────────────────────────────────
    offset = 0
    if n > ITEMS_SLOTS:
        extra = n - ITEMS_SLOTS
        offset = extra
        insert_at = ITEMS_END + 1
        ws.insert_rows(insert_at, extra)
        for i in range(extra):
            copy_row_style(ws, ITEMS_END, insert_at + i)

    # ── Fill item rows (rows 17 onward) ──────────────────────────────────────
    for i in range(n):
        r = ITEMS_START + i
        p = str(i)
        ws.cell(r, 1).value = i + 1
        ws.cell(r, 2).value = form.get(f'item_desc_{p}',   '')
        ws.cell(r, 3).value = form.get(f'item_model_{p}',  '')
        ws.cell(r, 5).value = form.get(f'item_color_{p}',  '')
        ws.cell(r, 6).value = form.get(f'item_size_{p}',   '')
        ws.cell(r, 7).value = form.get(f'item_unit_{p}',   'Set')
        try:    qty   = float(form.get(f'item_qty_{p}')   or 0)
        except: qty   = 0
        try:    price = float(form.get(f'item_price_{p}') or 0)
        except: price = 0
        ws.cell(r, 8).value  = qty
        ws.cell(r, 9).value  = price
        ws.cell(r, 10).value = f'=H{r}*I{r}'
        tax_raw  = form.get(f'item_tax_{p}', 'excl')
        if is_ar:
            tax_note = 'Incl. 16% GST شامل الضريبة' if tax_raw == 'incl' else 'Excl. Tax غير شامل الضريبة'
        else:
            tax_note = 'Incl. 16% GST 含16%消费税' if tax_raw == 'incl' else 'Excl. Tax 不含税'
        remarks  = (form.get(f'item_remarks_{p}') or '/').strip()
        ws.cell(r, 14).value = tax_note if remarks == '/' else f'{remarks} [{tax_note}]'

    # ── Delete unused item rows (so only filled rows show) ───────────────────
    deleted = 0
    if n < ITEMS_SLOTS:
        deleted = ITEMS_SLOTS - n
        delete_at = ITEMS_START + n
        ws.delete_rows(delete_at, deleted)

    shift = offset - deleted

    # ── Column widths: D(Picture)/K(Packing)/L(MOQ)/M(Lead) removed ──────
    for col in ('D', 'K', 'L', 'M'):
        ws.column_dimensions[col].hidden = True
    for col, w in {'A':4,'B':32,'C':11,'E':11,'F':11,'G':8,'H':7,
                   'I':12,'J':12,'N':12}.items():
        ws.column_dimensions[col].width = w

    # ── Summary rows (Subtotal=22, Freight=23, Grand=24, base) ───────────────
    last_item = ITEMS_START + n - 1
    sr = 22 + shift
    fr = 23 + shift
    gr = 24 + shift

    col_j = get_column_letter(10)

    ws.cell(sr, 10).value = f'=SUM({col_j}{ITEMS_START}:{col_j}{last_item})'
    try:    freight = float(form.get('freight') or 0)
    except: freight = 0
    ws.cell(fr, 10).value = freight
    ws.cell(gr, 10).value = f'={col_j}{sr}+{col_j}{fr}'
    ws.cell(gr, 14).value = currency

    # ── Sync validity into Remarks item 4 (base row 29) ──────────────────────
    remarks_row4 = 29 + shift
    cell_r4 = ws.cell(remarks_row4, 1)
    if cell_r4.value and '[  ]' in str(cell_r4.value):
        replacement = validity if validity else '___'
        cell_r4.value = str(cell_r4.value).replace('[  ]', replacement)

    # ── Company logo (A1:D2 area) ────────────────────────────────────────────
    selected_logo = form.get('selected_logo_name', '').strip()
    if selected_logo:
        logo_path = os.path.join(LOGO_DIR, selected_logo)
        if os.path.exists(logo_path):
            ws.cell(1, 1).value = None
            with open(logo_path, 'rb') as lf:
                insert_image(ws, lf, 'A1', w=150, h=70)

    # ── Update print area to reflect new last row ────────────────────────────
    last_row = 36 + shift
    ws.print_area = f'A1:N{last_row}'

    # ── Output ────────────────────────────────────────────────────────────────
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    inv_no   = (form.get('invoice_no') or 'PI').strip()
    customer = (form.get('customer_name') or '').strip()
    suffix   = '_AR' if is_ar else ''
    filename = f"PI_{inv_no}_{customer}{suffix}.xlsx".replace(' ', '_').replace('/', '-')

    return send_file(
        out,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename,
    )


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    d = date.today()
    today     = f"{d.year}.{d.month}.{d.day}"
    today_iso = d.isoformat()
    logos = logo_list_with_names()
    return render_template('index.html', today=today, today_iso=today_iso, logos=logos,
                           active_tab='pi-maker')


@app.route('/pl-maker')
def pl_maker():
    return render_template('placeholder.html',
                           title='PL Maker — 包装单制作',
                           active_tab='pl-maker')


@app.route('/quotation-maker')
def quotation_maker():
    return render_template('placeholder.html',
                           title='Quotation Maker — 报价单制作',
                           active_tab='quotation-maker')


@app.route('/jincheng-invoice')
def jincheng_invoice():
    return render_template('placeholder.html',
                           title='JINCHENG Invoice — 金城含税发票申请',
                           active_tab='jincheng-invoice')


@app.route('/logos')
def api_logos():
    return jsonify(logo_list_with_names())


@app.route('/upload-logo', methods=['POST'])
def upload_logo():
    f = request.files.get('logo_file')
    if not f or not f.filename:
        return jsonify({'error': 'No file provided'}), 400
    filename = secure_filename(f.filename)
    if not filename or os.path.splitext(filename)[1].lower() not in LOGO_EXTS:
        return jsonify({'error': 'Invalid file type'}), 400
    f.save(os.path.join(LOGO_DIR, filename))
    brand_name = request.form.get('brand_name', '').strip()
    if brand_name:
        names = load_logo_names()
        names[filename] = brand_name
        save_logo_names(names)
    display_name = brand_name or os.path.splitext(filename)[0].upper()
    return jsonify({'filename': filename, 'display_name': display_name})


@app.route('/logo/<path:filename>')
def serve_logo(filename):
    return send_from_directory(LOGO_DIR, filename)


@app.route('/customers')
def api_customers():
    return jsonify(load_customers())


@app.route('/save-customer', methods=['POST'])
def api_save_customer():
    data = request.get_json()
    if not data or not data.get('name', '').strip():
        return jsonify({'error': 'Customer name required'}), 400
    customers = load_customers()
    customer = {
        'id':      str(uuid.uuid4()),
        'name':    data.get('name',    '').strip(),
        'country': data.get('country', '').strip(),
        'address': data.get('address', '').strip(),
        'email':   data.get('email',   '').strip(),
        'tel':     data.get('tel',     '').strip(),
    }
    customers.append(customer)
    save_customers_file(customers)
    return jsonify(customer)


@app.route('/delete-customer/<cid>', methods=['DELETE'])
def api_delete_customer(cid):
    customers = [c for c in load_customers() if c.get('id') != cid]
    save_customers_file(customers)
    return jsonify({'ok': True})


@app.route('/products')
def api_products():
    products = load_products()
    project = request.args.get('project', '').strip()
    if project:
        products = [p for p in products if p.get('project', '') == project]
    return jsonify(products)


@app.route('/save-product', methods=['POST'])
def api_save_product():
    data = request.get_json() or {}
    if not data.get('desc', '').strip():
        return jsonify({'error': 'Product description required'}), 400
    products = load_products()
    product = {
        'id':        str(uuid.uuid4()),
        'project':   data.get('project',   '').strip(),
        'desc':      data.get('desc',      '').strip(),
        'model':     data.get('model',     '').strip(),
        'color':     data.get('color',     '').strip(),
        'size':      data.get('size',      '').strip(),
        'unit':      data.get('unit',      'Set').strip(),
        'price':     data.get('price',     '').strip(),
        'tax':       data.get('tax',       'excl').strip(),
        'packing':   data.get('packing',   'Carton').strip(),
        'moq':       data.get('moq',       '/').strip(),
        'lead_time': data.get('lead_time', '/').strip(),
        'remarks':   data.get('remarks',   '/').strip(),
    }
    products.append(product)
    save_products_file(products)
    return jsonify(product)


@app.route('/delete-product/<pid>', methods=['DELETE'])
def api_delete_product(pid):
    products = [p for p in load_products() if p.get('id') != pid]
    save_products_file(products)
    return jsonify({'ok': True})


@app.route('/translate', methods=['POST'])
def api_translate():
    data = request.get_json() or {}
    text = data.get('text', '').strip()
    if not text:
        return jsonify({'error': 'No text provided'}), 400
    try:
        params = urllib.parse.urlencode({'q': text, 'langpair': 'zh|en'})
        url = f'https://api.mymemory.translated.net/get?{params}'
        req = urllib.request.Request(url, headers={'User-Agent': 'PImaker/1.0'})
        with urllib.request.urlopen(req, timeout=10) as resp:
            result = json.loads(resp.read().decode())
        translated = result.get('responseData', {}).get('translatedText', '')
        if translated:
            return jsonify({'translated': translated})
        return jsonify({'error': 'No translation returned'}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/preview', methods=['POST'])
def preview():
    form  = request.form
    files = request.files
    n = max(1, int(form.get('item_count') or 1))

    items = []
    for i in range(n):
        p = str(i)
        try:    qty   = float(form.get(f'item_qty_{p}')   or 0)
        except: qty   = 0
        try:    price = float(form.get(f'item_price_{p}') or 0)
        except: price = 0
        tax_raw  = form.get(f'item_tax_{p}', 'excl')
        tax_note = 'Incl. 16% GST 含16%消费税' if tax_raw == 'incl' else 'Excl. Tax 不含税'
        remarks  = (form.get(f'item_remarks_{p}') or '/').strip()
        remark_combined = tax_note if remarks == '/' else f'{remarks} [{tax_note}]'
        items.append({
            'no':      i + 1,
            'desc':    form.get(f'item_desc_{p}',    ''),
            'model':   form.get(f'item_model_{p}',   ''),
            'color':   form.get(f'item_color_{p}',   ''),
            'size':    form.get(f'item_size_{p}',    ''),
            'unit':    form.get(f'item_unit_{p}',    'Set'),
            'qty':     qty,
            'price':   price,
            'amount':  qty * price,
            'remarks': remark_combined,
        })

    try:    freight = float(form.get('freight') or 0)
    except: freight = 0
    subtotal = sum(it['amount'] for it in items)
    grand    = subtotal + freight

    currency = form.get('currency') or 'JOD'

    date_raw = form.get('date', '')
    if date_raw and '-' in date_raw:
        try:
            d_obj = datetime.strptime(date_raw, '%Y-%m-%d')
            date_str = f"{d_obj.year}.{d_obj.month}.{d_obj.day}"
        except ValueError:
            date_str = date_raw
    else:
        date_str = date_raw

    prefix = (form.get('contact_prefix') or '').strip()
    person = (form.get('contact_person')  or '').strip()
    contact = f"{prefix} {person}".strip()

    port = (form.get('port_of_loading') or '').strip()
    dest = (form.get('destination') or '').strip()
    if port and dest:    shipping = f'{port} → {dest}'
    elif port:          shipping = port
    else:               shipping = dest


    # logo as base64 for inline embedding
    import base64
    logo_b64 = ''
    logo_mime = 'image/png'
    selected_logo = form.get('selected_logo_name', '').strip()
    if selected_logo:
        logo_path = os.path.join(LOGO_DIR, selected_logo)
        if os.path.exists(logo_path):
            ext = os.path.splitext(selected_logo)[1].lower()
            mime_map = {'.png': 'image/png', '.jpg': 'image/jpeg',
                        '.jpeg': 'image/jpeg', '.gif': 'image/gif', '.webp': 'image/webp'}
            logo_mime = mime_map.get(ext, 'image/png')
            with open(logo_path, 'rb') as lf:
                logo_b64 = base64.b64encode(lf.read()).decode()

    lang = form.get('lang', 'zh')
    is_ar = (lang == 'ar')

    # Rebuild items with language-aware tax notes
    for it in items:
        p = str(it['no'] - 1)  # 0-based index used in form field names
        tax_raw = form.get(f'item_tax_{p}', 'excl')
        if is_ar:
            tax_note = 'Incl. 16% GST شامل الضريبة' if tax_raw == 'incl' else 'Excl. Tax غير شامل الضريبة'
        else:
            tax_note = 'Incl. 16% GST 含16%消费税' if tax_raw == 'incl' else 'Excl. Tax 不含税'
        remarks = (form.get(f'item_remarks_{p}') or '/').strip()
        it['remarks'] = tax_note if remarks == '/' else f'{remarks} [{tax_note}]'

    ctx = {
        'company_name':    form.get('company_name', ''),
        'address':         form.get('address', ''),
        'city_country':    form.get('city_country', ''),
        'tel':             form.get('tel', ''),
        'email':           form.get('email', ''),
        'invoice_no':      form.get('invoice_no', ''),
        'date_str':        date_str,
        'customer_name':   form.get('customer_name', ''),
        'customer_address':form.get('customer_address', ''),
        'customer_email':  form.get('customer_email', ''),
        'country':         form.get('country', ''),
        'customer_tel':    form.get('customer_tel', ''),
        'contact':         contact,
        'inquiry_ref':     form.get('inquiry_ref', ''),
        'currency':        currency,
        'incoterms':       form.get('incoterms', ''),
        'payment_terms':   form.get('payment_terms', ''),
        'lead_time':       form.get('lead_time', ''),
        'validity':        form.get('quotation_validity', ''),
        'shipping':        shipping,
        'items':           items,
        'subtotal':        subtotal,
        'freight':         freight,
        'grand':           grand,
        'logo_b64':        logo_b64,
        'logo_mime':       logo_mime,
    }
    template = 'invoice_preview_ar.html' if is_ar else 'invoice_preview.html'
    return render_template(template, **ctx)


@app.route('/generate', methods=['POST'])
def generate():
    try:
        lang = request.form.get('lang', 'zh')
        return fill_template(request.form, request.files, lang=lang)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5000)
