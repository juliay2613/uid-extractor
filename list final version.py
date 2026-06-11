import sys
import os
import threading
import webbrowser
import json
import base64
import zipfile
import io
import urllib.request
import urllib.error
from xml.etree import ElementTree
from http.server import HTTPServer, BaseHTTPRequestHandler

XLSX_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def _col_letters(cell_ref):
    return ''.join(ch for ch in cell_ref if ch.isalpha())


def parse_xlsx(data):
    """Parse the first worksheet of an .xlsx file using only the standard library."""
    z = zipfile.ZipFile(io.BytesIO(data))

    shared_strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        root = ElementTree.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.findall(f'{XLSX_NS}si'):
            text = ''.join(t.text or '' for t in si.findall(f'.//{XLSX_NS}t'))
            shared_strings.append(text)

    sheet_root = ElementTree.fromstring(z.read('xl/worksheets/sheet1.xml'))

    rows = []
    for row in sheet_root.findall(f'.//{XLSX_NS}row'):
        row_data = {}
        for c in row.findall(f'{XLSX_NS}c'):
            ref = c.get('r')
            if not ref:
                continue
            col = _col_letters(ref)
            cell_type = c.get('t')
            value_el = c.find(f'{XLSX_NS}v')
            if value_el is not None:
                if cell_type == 's':
                    idx = int(value_el.text)
                    val = shared_strings[idx] if idx < len(shared_strings) else ''
                else:
                    val = value_el.text
            else:
                inline = c.find(f'{XLSX_NS}is')
                if inline is not None:
                    val = ''.join(t.text or '' for t in inline.findall(f'.//{XLSX_NS}t'))
                else:
                    val = None
            row_data[col] = val
        rows.append(row_data)

    if not rows:
        return []

    header_row = rows[0]
    headers = {col: name for col, name in header_row.items() if name}

    records = []
    for row in rows[1:]:
        record = {}
        for col, name in headers.items():
            record[name] = row.get(col)
        records.append(record)
    return records


def extract_reference_items(records):
    """One comparable item per order, using the order's first product and order total (since screenshots only show one product per order)."""
    items = []
    for record in records:
        store = record.get('店铺名称')
        order_no = record.get('订单号')
        group_key = record.get('分组键')
        name = record.get('商品1名称')
        if not name:
            continue
        items.append({
            'store': store or '',
            'orderNo': order_no or '',
            'groupKey': group_key or '',
            'name': name,
            'price': record.get('商品1划线原价') or '',
            'salePrice': record.get('商品1当前成交价') or '',
            'actualPaid': record.get('实际支付金额') or ''
        })
    return items


HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Product Price Extractor</title>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500&display=swap" rel="stylesheet">
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  :root {
    --bg: #0d0d0d; --surface: #161616; --surface2: #1f1f1f;
    --border: #2a2a2a; --border2: #383838; --text: #f0f0f0;
    --muted: #888; --accent: #ff2d55; --accent2: #ff6b6b;
    --success: #00d68f; --mono: 'IBM Plex Mono', monospace;
    --sans: 'IBM Plex Sans', sans-serif;
  }
  body { font-family: var(--sans); background: var(--bg); color: var(--text); min-height: 100vh; display: flex; flex-direction: column; align-items: center; padding: 40px 20px 80px; }
  header { width: 100%; max-width: 760px; margin-bottom: 40px; display: flex; align-items: flex-end; gap: 16px; border-bottom: 1px solid var(--border); padding-bottom: 20px; }
  .logo-mark { width: 36px; height: 36px; background: var(--accent); border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
  .logo-mark svg { width: 20px; height: 20px; stroke: white; fill: none; stroke-width: 2; }
  h1 { font-size: 18px; font-weight: 500; letter-spacing: -0.01em; line-height: 1; }
  h1 span { display: block; font-size: 12px; font-weight: 300; color: var(--muted); margin-top: 4px; letter-spacing: 0.05em; text-transform: uppercase; }
  .main { width: 100%; max-width: 760px; }
  .api-section { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 20px; margin-bottom: 20px; }
  .section-label { font-size: 11px; font-weight: 500; letter-spacing: 0.08em; text-transform: uppercase; color: var(--muted); margin-bottom: 12px; }
  .api-row { display: flex; gap: 8px; align-items: center; }
  .api-input { flex: 1; background: var(--surface2); border: 1px solid var(--border2); border-radius: 8px; padding: 10px 14px; font-family: var(--mono); font-size: 13px; color: var(--text); outline: none; transition: border-color 0.15s; }
  .api-input:focus { border-color: var(--accent); }
  .api-input::placeholder { color: #444; }
  .api-hint { font-size: 12px; color: var(--muted); margin-top: 8px; }
  .api-hint a { color: var(--accent2); text-decoration: none; }
  .drop-zone { background: var(--surface); border: 1px dashed var(--border2); border-radius: 12px; padding: 40px 20px; text-align: center; cursor: pointer; transition: border-color 0.15s, background 0.15s; margin-bottom: 16px; }
  .drop-zone:hover, .drop-zone.drag-over { border-color: var(--accent); background: #1a1212; }
  .drop-icon { width: 44px; height: 44px; margin: 0 auto 14px; border: 1px solid var(--border2); border-radius: 10px; display: flex; align-items: center; justify-content: center; }
  .drop-icon svg { width: 22px; height: 22px; stroke: var(--muted); fill: none; stroke-width: 1.5; }
  .drop-zone p { color: var(--muted); font-size: 14px; }
  .drop-zone p strong { color: var(--text); font-weight: 500; }
  .drop-zone small { font-size: 12px; color: #555; display: block; margin-top: 4px; }
  #thumbs { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 16px; }
  .thumb-wrap { position: relative; width: 80px; }
  .thumb-wrap img { width: 80px; height: 100px; object-fit: cover; border-radius: 8px; border: 1px solid var(--border2); display: block; }
  .thumb-name { font-size: 10px; color: var(--muted); margin-top: 4px; text-align: center; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .thumb-remove { position: absolute; top: -5px; right: -5px; width: 18px; height: 18px; background: var(--accent); border: none; border-radius: 50%; color: white; font-size: 11px; cursor: pointer; display: flex; align-items: center; justify-content: center; }
  .btn-row { display: flex; gap: 10px; margin-bottom: 24px; }
  .btn-primary { background: var(--accent); color: white; border: none; border-radius: 8px; padding: 11px 22px; font-family: var(--sans); font-size: 14px; font-weight: 500; cursor: pointer; transition: opacity 0.15s; display: flex; align-items: center; gap: 8px; }
  .btn-primary:hover { opacity: 0.88; }
  .btn-primary:disabled { opacity: 0.4; cursor: not-allowed; }
  .btn-ghost { background: transparent; color: var(--muted); border: 1px solid var(--border2); border-radius: 8px; padding: 11px 16px; font-family: var(--sans); font-size: 14px; cursor: pointer; }
  .btn-ghost:hover { color: var(--text); }
  .status-line { font-size: 13px; color: var(--muted); margin-bottom: 12px; }
  .status-line .err { color: #ff6b6b; }
  table.results-table { width: 100%; border-collapse: collapse; background: var(--surface); border: 1px solid var(--border); border-radius: 10px; overflow: hidden; margin-bottom: 16px; }
  table.results-table th, table.results-table td { text-align: left; padding: 10px 14px; font-size: 13px; border-bottom: 1px solid var(--border); }
  table.results-table th { font-size: 11px; font-weight: 500; letter-spacing: 0.06em; text-transform: uppercase; color: var(--muted); background: var(--surface2); }
  table.results-table tr:last-child td { border-bottom: none; }
  table.results-table td.price { font-family: var(--mono); }
  table.results-table td.sale-price { font-family: var(--mono); color: var(--success); }
  table.results-table td.file-cell { color: var(--muted); font-size: 11px; max-width: 140px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  table.results-table td.note-cell { color: var(--accent2); font-size: 12px; }
  table.results-table tr.row-flag { background: rgba(255, 45, 85, 0.08); }
  .copy-all-row { margin-top: 16px; padding-top: 16px; border-top: 1px solid var(--border); display: flex; align-items: center; justify-content: space-between; }
  .copy-all-row span { font-size: 13px; color: var(--muted); }
  footer { margin-top: 60px; font-size: 12px; color: #444; text-align: center; }
</style>
</head>
<body>
<header>
  <div class="logo-mark">
    <svg viewBox="0 0 24 24"><path d="M3 3h2l.4 2M7 13h10l4-8H5.4M7 13L5.4 5M7 13l-2.293 2.293A1 1 0 0 0 5.414 17H17M9 21a1 1 0 1 0 0-2 1 1 0 0 0 0 2zM20 21a1 1 0 1 0 0-2 1 1 0 0 0 0 2z"/></svg>
  </div>
  <h1>Product Price Extractor <span>Order List Screenshot Tool</span></h1>
</header>
<div class="main">
  <div class="api-section">
    <div class="section-label">Anthropic API Key</div>
    <div class="api-row">
      <input type="password" class="api-input" id="apiKey" placeholder="sk-ant-api03-..." autocomplete="off">
      <button class="btn-ghost" onclick="toggleKey()">Show</button>
    </div>
    <p class="api-hint">Your key is only used locally on your PC. Get one at <a href="https://console.anthropic.com/settings/keys" target="_blank">console.anthropic.com</a>.</p>
  </div>
  <div class="api-section">
    <div class="section-label">Reference Excel (optional)</div>
    <div class="api-row">
      <input type="file" class="api-input" id="excelInput" accept=".xlsx" style="padding:8px 14px;">
      <button class="btn-ghost" onclick="clearExcel()">Clear</button>
    </div>
    <p class="api-hint" id="excelStatus">Upload a reference .xlsx to verify extracted items against known order data.</p>
  </div>
  <div class="drop-zone" id="dropZone" onclick="document.getElementById('fileInput').click()">
    <div class="drop-icon">
      <svg viewBox="0 0 24 24"><polyline points="16 16 12 12 8 16"/><line x1="12" y1="12" x2="12" y2="21"/><path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/></svg>
    </div>
    <p><strong>Drop order list screenshots here</strong> or click to browse</p>
    <small>PNG, JPG — multiple files OK</small>
  </div>
  <input type="file" id="fileInput" accept="image/*" multiple style="display:none">
  <div id="thumbs"></div>
  <div class="btn-row" id="btnRow" style="display:none">
    <button class="btn-primary" id="extractBtn" onclick="extractAll()">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
      Extract Products
    </button>
    <button class="btn-ghost" id="retryBtn" onclick="retryUnidentified()" style="display:none">Re-extract unidentified</button>
    <button class="btn-ghost" onclick="clearAll()">Clear all</button>
  </div>
  <div id="status"></div>
  <div id="results"></div>
  <div id="copyAllRow" class="copy-all-row" style="display:none">
    <span id="copyAllLabel">0 items found</span>
    <button class="btn-primary" onclick="copyAllCSV()" style="padding:8px 16px;font-size:13px;">Copy all as CSV</button>
    <button class="btn-primary" onclick="downloadExcel()" style="padding:8px 16px;font-size:13px;">Download as Excel</button>
  </div>
</div>
<footer>Running locally on your PC — no data leaves your machine except to Anthropic's API.</footer>

<script>
const files = [];
let extractedItems = [];
let excelItems = [];
let unmatchedExcelItems = [];
let excelMatchPool = [];
let fileEmpty = {};
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const excelInput = document.getElementById('excelInput');
const excelStatus = document.getElementById('excelStatus');
const thumbsEl = document.getElementById('thumbs');
const resultsEl = document.getElementById('results');
const statusEl = document.getElementById('status');
const btnRow = document.getElementById('btnRow');
const retryBtn = document.getElementById('retryBtn');
const copyAllRow = document.getElementById('copyAllRow');

dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => { e.preventDefault(); dropZone.classList.remove('drag-over'); addFiles(e.dataTransfer.files); });
fileInput.addEventListener('change', e => { addFiles(e.target.files); fileInput.value = ''; });

excelInput.addEventListener('change', async e => {
  const file = e.target.files[0];
  if (!file) return;
  excelStatus.textContent = `Reading ${file.name}…`;
  try {
    const buf = await file.arrayBuffer();
    let binary = '';
    const bytes = new Uint8Array(buf);
    for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
    const b64 = btoa(binary);
    const resp = await fetch('/api/parse-excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fileData: b64 })
    });
    const data = await resp.json();
    if (data.error) throw new Error(data.error);
    excelItems = data.items || [];
    excelStatus.textContent = `Loaded ${excelItems.length} reference item(s) from ${file.name}.`;
    if (extractedItems.length > 0) renderTable();
  } catch (err) {
    excelStatus.textContent = `Error reading excel: ${err.message}`;
    excelItems = [];
  }
});

function clearExcel() {
  excelInput.value = '';
  excelItems = [];
  unmatchedExcelItems = [];
  excelStatus.textContent = 'Upload a reference .xlsx to verify extracted items against known order data.';
  if (extractedItems.length > 0) renderTable();
}

function addFiles(list) {
  Array.from(list).forEach(f => {
    if (!f.type.startsWith('image/')) return;
    const idx = files.length;
    files.push(f);
    const url = URL.createObjectURL(f);
    const wrap = document.createElement('div');
    wrap.className = 'thumb-wrap';
    wrap.id = 'tw-' + idx;
    wrap.innerHTML = `<button class="thumb-remove" onclick="removeFile(${idx})">×</button><img src="${url}" alt=""><div class="thumb-name">${f.name}</div>`;
    thumbsEl.appendChild(wrap);
  });
  updateUI();
}

function removeFile(idx) { files.splice(idx, 1); rebuildThumbs(); updateUI(); }

function rebuildThumbs() {
  thumbsEl.innerHTML = '';
  files.forEach((f, i) => {
    const url = URL.createObjectURL(f);
    const wrap = document.createElement('div');
    wrap.className = 'thumb-wrap';
    wrap.id = 'tw-' + i;
    wrap.innerHTML = `<button class="thumb-remove" onclick="removeFile(${i})">×</button><img src="${url}" alt=""><div class="thumb-name">${f.name}</div>`;
    thumbsEl.appendChild(wrap);
  });
}

function updateUI() { btnRow.style.display = files.length ? 'flex' : 'none'; }
function clearAll() {
  files.length = 0; thumbsEl.innerHTML = ''; resultsEl.innerHTML = ''; statusEl.innerHTML = '';
  copyAllRow.style.display = 'none'; extractedItems = []; unmatchedExcelItems = []; fileEmpty = {};
  retryBtn.style.display = 'none'; updateUI();
}
function toggleKey() {
  const inp = document.getElementById('apiKey');
  const btn = event.target;
  if (inp.type === 'password') { inp.type = 'text'; btn.textContent = 'Hide'; }
  else { inp.type = 'password'; btn.textContent = 'Show'; }
}

async function toBase64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(',')[1]);
    r.onerror = () => rej(new Error('Read failed'));
    r.readAsDataURL(file);
  });
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function normalizePrice(val) {
  const digits = String(val ?? '').replace(/[^0-9]/g, '');
  return digits ? parseInt(digits, 10) : null;
}

function normalizeStore(val) {
  return (val ?? '').trim().toLowerCase().replace(/[.…]+$/, '').trim();
}

function storesMatch(a, b) {
  if (!a || !b) return false;
  if (a === b) return true;
  return a.startsWith(b) || b.startsWith(a);
}

function resolveGroupKey(filename) {
  const base = String(filename ?? '').replace(/\.[^.]+$/, '').toLowerCase();
  for (const it of excelItems) {
    const key = (it.groupKey || '').toLowerCase();
    if (key && base.startsWith(key)) return it.groupKey;
  }
  return null;
}

function formatPrice(val) {
  const n = normalizePrice(val);
  return n === null ? '' : String(n);
}

function tokenize(s) {
  return new Set(String(s ?? '').toLowerCase().match(/[a-z0-9一-鿿]+/g) || []);
}

function nameOverlap(a, b) {
  const ta = tokenize(a), tb = tokenize(b);
  let count = 0;
  for (const t of ta) if (tb.has(t)) count++;
  return count;
}

const MATCH_PASSES = [
  {
    note: '',
    test: (ref, store, price, salePrice, item, orderTotal) => storesMatch(ref._store, store)
      && nameOverlap(ref.name, item.name) >= 1
      && ref._salePrice === salePrice
      && ref._actualPaid === orderTotal
  },
  {
    note: 'Order total differs 实付总额不一致',
    test: (ref, store, price, salePrice, item) => storesMatch(ref._store, store)
      && nameOverlap(ref.name, item.name) >= 1
      && ref._salePrice === salePrice
  },
  {
    note: 'Sale price differs 实付价不一致',
    test: (ref, store, price, salePrice, item, orderTotal) => storesMatch(ref._store, store)
      && nameOverlap(ref.name, item.name) >= 1
      && ref._actualPaid === orderTotal
  },
  {
    note: 'Matched by store and product name only, prices differ, please verify 仅按店铺和商品名称匹配,价格不符,请核对',
    test: (ref, store, price, salePrice, item) => storesMatch(ref._store, store)
      && nameOverlap(ref.name, item.name) >= 1
  },
  {
    note: 'Matched by product name and price only, please verify 仅按商品名称和价格匹配,请核对',
    test: (ref, store, price, salePrice, item, orderTotal) => nameOverlap(ref.name, item.name) >= 1
      && (ref._salePrice === salePrice || ref._actualPaid === orderTotal)
  },
  {
    note: 'Matched by store name only, please verify 仅按店铺名称匹配,请核对',
    test: (ref, store, price, salePrice, item) => storesMatch(ref._store, store)
  }
];

function matchAgainstExcel() {
  const pool = excelItems.map(it => ({
    ...it,
    _store: normalizeStore(it.store),
    _price: normalizePrice(it.price),
    _salePrice: normalizePrice(it.salePrice),
    _actualPaid: normalizePrice(it.actualPaid),
    _matchedItem: null
  }));
  extractedItems.forEach(item => {
    item.matched = false;
    item.ambiguous = false;
    item.matchedRef = null;
    item.note = '';
  });
  MATCH_PASSES.forEach(pass => {
    extractedItems.forEach(item => {
      if (item.matched) return;
      const price = normalizePrice(item.price);
      const salePrice = normalizePrice(item.salePrice);
      const orderTotal = normalizePrice(item.orderTotal);
      const store = normalizeStore(item.store);
      let candidates = pool.filter(ref => !ref._used && pass.test(ref, store, price, salePrice, item, orderTotal));
      if (candidates.length > 1) {
        const scores = candidates.map(c => nameOverlap(c.name, item.name));
        const maxScore = Math.max(...scores);
        const best = candidates.filter((c, i) => scores[i] === maxScore);
        if (best.length === 1) candidates = best;
      }
      if (candidates.length > 0) {
        candidates[0]._used = true;
        candidates[0]._matchedItem = item;
        item.matched = true;
        item.ambiguous = candidates.length > 1;
        item.matchedRef = candidates[0];
        item.note = pass.note;
      }
    });
  });
  extractedItems.forEach(item => {
    if (item.matched) return;
    const salePrice = normalizePrice(item.salePrice);
    const orderTotal = normalizePrice(item.orderTotal);
    const isDuplicate = extractedItems.some(other => {
      if (other === item || !other.matched) return false;
      return normalizePrice(other.salePrice) === salePrice
        && normalizePrice(other.orderTotal) === orderTotal
        && nameOverlap(other.name, item.name) >= 1;
    });
    item.duplicate = isDuplicate;
  });
  pool.filter(ref => !ref._used).forEach(ref => {
    let storeBest = null;
    let storeBestScore = -1;
    let best = null;
    let bestScore = 0;
    extractedItems.forEach(item => {
      const score = nameOverlap(ref.name, item.name);
      if (storesMatch(ref._store, normalizeStore(item.store)) && score > storeBestScore) {
        storeBestScore = score;
        storeBest = item;
      }
      if (score > bestScore) { bestScore = score; best = item; }
    });
    if (storeBest) {
      ref._guessItem = storeBest;
    } else if (best && bestScore >= 1) {
      ref._guessItem = best;
    }
  });
  unmatchedExcelItems = pool.filter(ref => !ref._used);
  excelMatchPool = pool;
}

function buildNote(item) {
  const notes = [];
  if (normalizePrice(item.price) === null || normalizePrice(item.salePrice) === null) {
    notes.push('Price could not be read');
  }
  if (excelItems.length > 0) {
    if (!item.matched) notes.push('No matching row in excel');
    else if (item.ambiguous) notes.push('Multiple excel rows match — variant unclear');
  }
  return notes.join('; ');
}

function renderTable() {
  if (extractedItems.length === 0) {
    resultsEl.innerHTML = '';
    copyAllRow.style.display = 'none';
    return;
  }
  if (excelItems.length > 0) matchAgainstExcel();

  const matchHeader = excelItems.length > 0 ? '<th>Match</th>' : '';
  let rows = extractedItems.map(item => {
    const matchCell = excelItems.length > 0
      ? `<td>${item.matched ? '<span style="color:var(--success)">&#10003; Matched</span>' : '<span style="color:var(--accent2)">&#10007; Not in excel</span>'}</td>`
      : '';
    const note = buildNote(item);
    item.note = note;
    const rowClass = note ? ' class="row-flag"' : '';
    const groupKey = resolveGroupKey(item.file) || item.file;
    return `
    <tr${rowClass}>
      <td class="file-cell">${escapeHtml(groupKey)}</td>
      <td>${escapeHtml(item.store)}</td>
      <td>${escapeHtml(item.name)}</td>
      <td class="price">${escapeHtml(formatPrice(item.price))}</td>
      <td class="sale-price">${escapeHtml(formatPrice(item.salePrice))}</td>
      ${matchCell}
      <td class="note-cell">${escapeHtml(note)}</td>
    </tr>`;
  }).join('');

  let unmatchedHtml = '';
  if (excelItems.length > 0 && unmatchedExcelItems.length > 0) {
    const unmatchedRows = unmatchedExcelItems.map(ref => `
      <tr>
        <td class="file-cell">${escapeHtml(ref.orderNo || '')}</td>
        <td>${escapeHtml(ref.store)}</td>
        <td>${escapeHtml(ref.name)}</td>
        <td class="price">${escapeHtml(ref.price)}</td>
        <td class="sale-price">${escapeHtml(ref.salePrice)}</td>
      </tr>`).join('');
    unmatchedHtml = `
      <p class="status-line" style="margin-top:16px">Excel items not found in screenshots (${unmatchedExcelItems.length}):</p>
      <table class="results-table">
        <thead><tr><th>Order No.</th><th>Store</th><th>Product Name</th><th>Price</th><th>Sale Price</th></tr></thead>
        <tbody>${unmatchedRows}</tbody>
      </table>`;
  }

  resultsEl.innerHTML = `
    <table class="results-table">
      <thead><tr><th>Group Key</th><th>Store</th><th>Product Name</th><th>Price</th><th>Sale Price</th>${matchHeader}<th>Note</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>${unmatchedHtml}`;
  copyAllRow.style.display = 'flex';
  document.getElementById('copyAllLabel').textContent = `${extractedItems.length} item${extractedItems.length > 1 ? 's' : ''} found`;
}

async function extractFile(i) {
  const apiKey = document.getElementById('apiKey').value.trim();
  statusEl.innerHTML = `<p class="status-line">Analyzing ${escapeHtml(files[i].name)} (${i + 1}/${files.length})…</p>`;

  try {
    const b64 = await toBase64(files[i]);
    const resp = await fetch('/api/extract', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ apiKey, imageData: b64, mediaType: files[i].type || 'image/jpeg' })
    });

    const data = await resp.json();
    if (data.error) throw new Error(typeof data.error === 'object' ? data.error.message : data.error);

    const raw = data.content?.find(c => c.type === 'text')?.text || '';
    let parsed;
    try { parsed = JSON.parse(raw.replace(/```json|```/g, '').trim()); } catch(e) { parsed = null; }

    const items = Array.isArray(parsed?.items) ? parsed.items : [];
    let lastStore = '';
    let lastOrderTotal = '';
    let added = 0;
    items.forEach(it => {
      const store = it.store || lastStore;
      if (it.store) lastStore = it.store;
      const orderTotal = it.order_total || lastOrderTotal;
      if (it.order_total) lastOrderTotal = it.order_total;
      const item = {
        file: files[i].name,
        store: store,
        name: it.name || '',
        price: it.price ?? '',
        salePrice: it.sale_price ?? '',
        orderTotal: orderTotal
      };
      const key = `${item.store}|${item.name}|${item.price}|${item.salePrice}`;
      if (!extractedItems.some(e => `${e.store}|${e.name}|${e.price}|${e.salePrice}` === key)) {
        extractedItems.push(item);
        added++;
      }
    });
    fileEmpty[files[i].name] = added === 0;
    renderTable();
  } catch(err) {
    statusEl.innerHTML = `<p class="status-line"><span class="err">Error on ${escapeHtml(files[i].name)}: ${escapeHtml(err.message)}</span></p>`;
    fileEmpty[files[i].name] = true;
  }
}

async function extractAll() {
  const apiKey = document.getElementById('apiKey').value.trim();
  if (!apiKey) { alert('Please enter your Anthropic API key first.'); return; }
  resultsEl.innerHTML = '';
  statusEl.innerHTML = '';
  extractedItems = [];
  fileEmpty = {};
  copyAllRow.style.display = 'none';
  document.getElementById('extractBtn').disabled = true;
  retryBtn.style.display = 'none';

  for (let i = 0; i < files.length; i++) {
    await extractFile(i);
  }

  document.getElementById('extractBtn').disabled = false;
  if (extractedItems.length > 0) {
    statusEl.innerHTML = `<p class="status-line">Done. ${extractedItems.length} item(s) found across ${files.length} file(s).</p>`;
  } else if (!statusEl.querySelector('.err')) {
    statusEl.innerHTML = `<p class="status-line">No products detected.</p>`;
  }
  updateRetryButton();
}

function fileNeedsRetry(filename) {
  if (fileEmpty[filename]) return true;
  if (excelItems.length === 0) return false;
  const items = extractedItems.filter(it => it.file === filename);
  return items.length > 0 && items.every(it => !it.matched);
}

function updateRetryButton() {
  const failedCount = files.filter(f => fileNeedsRetry(f.name)).length;
  if (failedCount > 0) {
    retryBtn.textContent = `Re-extract unidentified (${failedCount})`;
    retryBtn.style.display = '';
  } else {
    retryBtn.style.display = 'none';
  }
}

async function retryUnidentified() {
  const apiKey = document.getElementById('apiKey').value.trim();
  if (!apiKey) { alert('Please enter your Anthropic API key first.'); return; }
  const indices = files.map((f, i) => i).filter(i => fileNeedsRetry(files[i].name));
  if (indices.length === 0) return;
  document.getElementById('extractBtn').disabled = true;
  retryBtn.disabled = true;

  for (const i of indices) {
    await extractFile(i);
  }

  document.getElementById('extractBtn').disabled = false;
  retryBtn.disabled = false;
  statusEl.innerHTML = `<p class="status-line">Done. ${extractedItems.length} item(s) found across ${files.length} file(s).</p>`;
  updateRetryButton();
}

function csvEscape(val) {
  const s = String(val ?? '');
  if (/[",\\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function copyAllCSV() {
  const hasMatch = excelItems.length > 0;
  const header = ['Group Key', 'Store', 'Product Name', 'Price', 'Sale Price', ...(hasMatch ? ['Match'] : []), 'Note'].join(',');
  const rows = extractedItems.map(it => {
    const groupKey = resolveGroupKey(it.file) || it.file;
    const cols = [groupKey, it.store, it.name, formatPrice(it.price), formatPrice(it.salePrice)];
    if (hasMatch) cols.push(it.matched ? 'Matched' : 'Not in excel');
    cols.push(it.note || '');
    return cols.map(csvEscape).join(',');
  });
  const csv = [header, ...rows].join('\\n');
  navigator.clipboard.writeText(csv).then(() => {
    const btn = event.target;
    const orig = btn.innerHTML;
    btn.innerHTML = '✓ Copied!';
    setTimeout(() => { btn.innerHTML = orig; }, 1500);
  });
}

function xmlEscape(val) {
  return String(val ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
}

function xmlCell(val) {
  const s = String(val ?? '');
  if (s !== '' && /^-?[0-9]+(\\.[0-9]+)?$/.test(s)) {
    return `<Cell><Data ss:Type="Number">${xmlEscape(s)}</Data></Cell>`;
  }
  return `<Cell><Data ss:Type="String">${xmlEscape(s)}</Data></Cell>`;
}

function xmlRow(cells, flagged) {
  let styleId = '';
  if (flagged === true) styleId = 'Flag';
  else if (flagged === 'yellow') styleId = 'FlagYellow';
  const style = styleId ? ` ss:StyleID="${styleId}"` : '';
  return `<Row${style}>${cells.map(xmlCell).join('')}</Row>`;
}

function displayStore(store) {
  return store ? store : 'No Shop Name 无店铺名称';
}

function buildWorkbookXml(rows) {
  return `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal"/>
  <Style ss:ID="Flag"><Interior ss:Color="#FFE6E6" ss:Pattern="Solid"/></Style>
  <Style ss:ID="FlagYellow"><Interior ss:Color="#FFF9C4" ss:Pattern="Solid"/></Style>
 </Styles>
 <Worksheet ss:Name="Comparison">
  <Table>
${rows.join('\\n')}
  </Table>
 </Worksheet>
</Workbook>`;
}

function triggerDownload(xml) {
  const blob = new Blob([xml], { type: 'application/vnd.ms-excel' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'order_comparison.xls';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function downloadExcel() {
  if (excelItems.length > 0) {
    matchAgainstExcel();
    downloadExcelByOrder();
  } else {
    downloadExcelFlat();
  }
}

function downloadExcelFlat() {
  const headers = [
    'Group Key 分组键', 'Store 店铺', 'Product Name 商品名称',
    'Screenshot Price 截图原价', 'Screenshot Sale Price 截图实付价'
  ];

  const rows = [xmlRow(headers, false)];

  extractedItems.forEach(item => {
    const groupKey = resolveGroupKey(item.file) || item.file;
    rows.push(xmlRow([
      groupKey, displayStore(item.store), item.name,
      formatPrice(item.price), formatPrice(item.salePrice)
    ], !item.store));
  });

  triggerDownload(buildWorkbookXml(rows));
}

function downloadExcelByOrder() {
  const headers = [
    'No.', 'Group Key 分组键', 'Store 店铺',
    'Excel Product 表格商品', 'Excel Sale 表格实付价', 'Excel Order Total 表格实付总额',
    'Screenshot Product 截图商品', 'Screenshot Sale 截图实付价', 'Screenshot Order Total 截图实付总额',
    'Matched 是否匹配', 'Note 备注'
  ];

  const rows = [xmlRow(headers, false)];
  let counter = 0;

  excelMatchPool.forEach(ref => {
    const item = ref._matchedItem;
    let note = '';
    let flagged = false;
    let matched = 0;
    let display = item;
    if (!item) {
      if (ref._guessItem) {
        display = ref._guessItem;
        matched = 1;
        flagged = true;
        note = 'Possible match by product name/store, please double check 按商品名称/店铺可能匹配,请核对';
      } else {
        flagged = 'yellow';
        note = 'Not found in screenshots 截图中未找到';
      }
    } else {
      matched = 1;
      if (item.ambiguous) {
        note = 'Multiple possible matches 多个可能匹配';
        flagged = true;
      } else if (item.note) {
        note = item.note;
        flagged = true;
      }
    }
    counter++;
    rows.push(xmlRow([
      counter, ref.groupKey || '', ref.store || '',
      ref.name || '', formatPrice(ref.salePrice), formatPrice(ref.actualPaid),
      display ? display.name : '', display ? formatPrice(display.salePrice) : '', display ? formatPrice(display.orderTotal) : '',
      matched, note
    ], flagged));
  });

  extractedItems.filter(item => !item.matched && !item.duplicate).forEach(item => {
    const groupKey = resolveGroupKey(item.file) || item.file;
    counter++;
    rows.push(xmlRow([
      counter, groupKey, displayStore(item.store),
      '', '', '',
      item.name, formatPrice(item.salePrice), formatPrice(item.orderTotal),
      0, 'Not found in excel 表格中未找到'
    ], 'yellow'));
  });

  triggerDownload(buildWorkbookXml(rows));
}
</script>
</body>
</html>"""

class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # suppress console logs

    def do_GET(self):
        if self.path == '/' or self.path == '/index.html':
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(HTML.encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if self.path == '/api/extract':
            length = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(length))

            api_key = body.get('apiKey', '')
            image_data = body.get('imageData', '')
            media_type = body.get('mediaType', 'image/jpeg')

            payload = json.dumps({
                "model": "claude-sonnet-4-6",
                "max_tokens": 2500,
                "system": "You are a precise data extractor. The image is a screenshot of an order list / product list on an e-commerce product page. The screenshot may contain MULTIPLE separate order cards stacked vertically, each with its own store/seller name, product, price, and 'Total X produk: RpY' line - you MUST extract EVERY single order card visible in the image, from top to bottom, no matter how many there are (could be 1, 2, 3, 4 or more). Do not stop after the first one and do not skip any order card just because it looks similar to another. Each order card is a separate entry in the 'items' array, even if the store name or product looks similar to another card. For EACH product/item visible, extract: store (the seller/shop name shown for that order, or null if not visible), name (the product title/name as shown), price (the original/listed price as shown, including currency symbol, as a string), sale_price (the actual price after any discount/sale, including currency symbol, as a string; if there is no separate sale price, use the same value as price), order_total (the order/parcel total amount shown for this item's order, including currency symbol, as a string, or null if not visible). IMPORTANT: price and sale_price must be the price shown directly next to or below that specific product (the per-item price), NOT the order_total, which can include shipping fees, service fees, or voucher discounts that apply to the whole order rather than this single item; order_total should still be captured separately in its own field. CRITICAL for order_total: the order summary line is often shown as 'Total X produk: RpY' or 'Total X Produk Rp Y', where X is the NUMBER OF PRODUCTS in the order (a small integer like 1-20) and Y is the actual Rupiah amount. order_total must be ONLY the amount Y (the part after 'produk:' / 'Produk', i.e. the Rp amount), and must NOT include the leading product-count digit X. For example, if the screenshot shows 'Total 7 produk: Rp147.720', order_total must be 'Rp147.720' (i.e. 147720), NOT 'Rp7147.720' (i.e. 7147720) - the '7' is the product count and is a completely separate number from the total amount, do not prepend or merge it into order_total. Return ONLY a JSON object with a single key 'items', an array of objects each with keys: store, name, price, sale_price, order_total. No markdown, no extra text. If no products are found, return {\"items\":[]}.",
                "messages": [{
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": image_data}},
                        {"type": "text", "text": "Extract all products with their store/seller name, product name, original price, sale price (after discount), and order total from this order list screenshot. If this screenshot shows multiple separate order cards, extract every single one of them - do not skip any."}
                    ]
                }]
            }).encode('utf-8')

            req = urllib.request.Request(
                'https://api.anthropic.com/v1/messages',
                data=payload,
                headers={
                    'Content-Type': 'application/json',
                    'x-api-key': api_key,
                    'anthropic-version': '2023-06-01'
                },
                method='POST'
            )

            try:
                with urllib.request.urlopen(req, timeout=30) as res:
                    result = res.read()
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(result)
            except urllib.error.HTTPError as e:
                err_body = e.read()
                self.send_response(e.code)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(err_body)
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode())
        elif self.path == '/api/parse-excel':
            length = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(length))
            file_data = body.get('fileData', '')

            try:
                raw = base64.b64decode(file_data)
                records = parse_xlsx(raw)
                items = extract_reference_items(records)
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"items": items}).encode())
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode())
        else:
            self.send_response(404)
            self.end_headers()


def find_free_port():
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        return s.getsockname()[1]

def run():
    port = find_free_port()
    server = HTTPServer(('127.0.0.1', port), Handler)
    url = f'http://127.0.0.1:{port}'

    threading.Timer(0.5, lambda: webbrowser.open(url)).start()

    print(f"Product Price Extractor running at {url}")
    print("Close this window to stop the app.")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass

if __name__ == '__main__':
    run()
