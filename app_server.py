import sys
import os
import threading
import webbrowser
import json
import base64
import urllib.request
import urllib.error
from http.server import HTTPServer, BaseHTTPRequestHandler

HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>TikTok UID Extractor</title>
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
  header { width: 100%; max-width: 640px; margin-bottom: 40px; display: flex; align-items: flex-end; gap: 16px; border-bottom: 1px solid var(--border); padding-bottom: 20px; }
  .logo-mark { width: 36px; height: 36px; background: var(--accent); border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
  .logo-mark svg { width: 20px; height: 20px; fill: white; }
  h1 { font-size: 18px; font-weight: 500; letter-spacing: -0.01em; line-height: 1; }
  h1 span { display: block; font-size: 12px; font-weight: 300; color: var(--muted); margin-top: 4px; letter-spacing: 0.05em; text-transform: uppercase; }
  .main { width: 100%; max-width: 640px; }
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
  .result-card { background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 16px 18px; margin-bottom: 10px; animation: fadeIn 0.25s ease; }
  @keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: none; } }
  .result-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }
  .result-filename { font-size: 12px; color: var(--muted); }
  .badge { font-size: 11px; font-weight: 500; padding: 3px 8px; border-radius: 5px; letter-spacing: 0.04em; }
  .badge-ok { background: #0a2e20; color: var(--success); }
  .badge-err { background: #2e0a10; color: #ff6b6b; }
  .badge-loading { background: var(--surface2); color: var(--muted); }
  .uid-row { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 10px 12px; background: var(--surface2); border-radius: 8px; margin-bottom: 6px; }
  .uid-label { font-size: 11px; color: var(--muted); margin-bottom: 3px; letter-spacing: 0.04em; text-transform: uppercase; }
  .uid-value { font-family: var(--mono); font-size: 16px; font-weight: 500; color: var(--text); letter-spacing: 0.02em; }
  .copy-btn { background: transparent; border: 1px solid var(--border2); color: var(--muted); border-radius: 6px; padding: 5px 10px; font-size: 12px; cursor: pointer; flex-shrink: 0; white-space: nowrap; }
  .copy-btn:hover { color: var(--text); }
  .copy-btn.copied { color: var(--success); border-color: var(--success); }
  .spinner { display: inline-block; width: 14px; height: 14px; border: 2px solid var(--border2); border-top-color: var(--muted); border-radius: 50%; animation: spin 0.7s linear infinite; vertical-align: -2px; margin-right: 6px; }
  @keyframes spin { to { transform: rotate(360deg); } }
  .copy-all-row { margin-top: 16px; padding-top: 16px; border-top: 1px solid var(--border); display: flex; align-items: center; justify-content: space-between; }
  .copy-all-row span { font-size: 13px; color: var(--muted); }
  footer { margin-top: 60px; font-size: 12px; color: #444; text-align: center; }
</style>
</head>
<body>
<header>
  <div class="logo-mark">
    <svg viewBox="0 0 24 24"><path d="M19.59 6.69a4.83 4.83 0 0 1-3.77-4.25V2h-3.45v13.67a2.89 2.89 0 0 1-2.88 2.5 2.89 2.89 0 0 1-2.89-2.89 2.89 2.89 0 0 1 2.89-2.89c.28 0 .54.04.79.1V9.01a6.33 6.33 0 0 0-.79-.05 6.34 6.34 0 0 0-6.34 6.34 6.34 6.34 0 0 0 6.34 6.34 6.34 6.34 0 0 0 6.33-6.34l-.03-8.27a8.2 8.2 0 0 0 4.79 1.53V5.12a4.85 4.85 0 0 1-1-.43z"/></svg>
  </div>
  <h1>UID Extractor <span>TikTok Settings Screenshot Tool</span></h1>
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
  <div class="drop-zone" id="dropZone" onclick="document.getElementById('fileInput').click()">
    <div class="drop-icon">
      <svg viewBox="0 0 24 24"><polyline points="16 16 12 12 8 16"/><line x1="12" y1="12" x2="12" y2="21"/><path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/></svg>
    </div>
    <p><strong>Drop screenshots here</strong> or click to browse</p>
    <small>PNG, JPG — multiple files OK</small>
  </div>
  <input type="file" id="fileInput" accept="image/*" multiple style="display:none">
  <div id="thumbs"></div>
  <div class="btn-row" id="btnRow" style="display:none">
    <button class="btn-primary" id="extractBtn" onclick="extractAll()">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
      Extract UIDs
    </button>
    <button class="btn-ghost" onclick="clearAll()">Clear all</button>
  </div>
  <div id="results"></div>
  <div id="copyAllRow" class="copy-all-row" style="display:none">
    <span id="copyAllLabel">0 UIDs found</span>
    <button class="btn-primary" onclick="copyAllUIDs()" style="padding:8px 16px;font-size:13px;">Copy all UIDs</button>
  </div>
</div>
<footer>Running locally on your PC — no data leaves your machine except to Anthropic's API.</footer>

<script>
const files = [];
let extractedUIDs = [];
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const thumbsEl = document.getElementById('thumbs');
const resultsEl = document.getElementById('results');
const btnRow = document.getElementById('btnRow');
const copyAllRow = document.getElementById('copyAllRow');

dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => { e.preventDefault(); dropZone.classList.remove('drag-over'); addFiles(e.dataTransfer.files); });
fileInput.addEventListener('change', e => { addFiles(e.target.files); fileInput.value = ''; });

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
function clearAll() { files.length = 0; thumbsEl.innerHTML = ''; resultsEl.innerHTML = ''; copyAllRow.style.display = 'none'; extractedUIDs = []; updateUI(); }
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

async function extractAll() {
  const apiKey = document.getElementById('apiKey').value.trim();
  if (!apiKey) { alert('Please enter your Anthropic API key first.'); return; }
  resultsEl.innerHTML = '';
  extractedUIDs = [];
  copyAllRow.style.display = 'none';
  document.getElementById('extractBtn').disabled = true;

  for (let i = 0; i < files.length; i++) {
    const card = document.createElement('div');
    card.className = 'result-card';
    card.id = 'rc-' + i;
    card.innerHTML = `<div class="result-header"><span class="result-filename">${files[i].name}</span><span class="badge badge-loading"><span class="spinner"></span>Analyzing…</span></div>`;
    resultsEl.appendChild(card);

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

      if (parsed && parsed.uid) {
        extractedUIDs.push({ file: files[i].name, uid: parsed.uid });
        card.innerHTML = `
          <div class="result-header"><span class="result-filename">${files[i].name}</span><span class="badge badge-ok">✓ Found</span></div>
          <div class="uid-row"><div><div class="uid-label">${parsed.source || 'UID'}</div><div class="uid-value" id="uid-val-${i}">${parsed.uid}</div></div><button class="copy-btn" onclick="copyVal('uid-val-${i}', this)">Copy</button></div>
          ${parsed.did ? `<div class="uid-row"><div><div class="uid-label">DID</div><div class="uid-value" id="did-val-${i}">${parsed.did}</div></div><button class="copy-btn" onclick="copyVal('did-val-${i}', this)">Copy</button></div>` : ''}`;
      } else {
        card.innerHTML = `<div class="result-header"><span class="result-filename">${files[i].name}</span><span class="badge badge-err">Not found</span></div><p style="font-size:13px;color:#ff6b6b;margin-top:4px">No UID detected in this image.</p>`;
      }
    } catch(err) {
      card.innerHTML = `<div class="result-header"><span class="result-filename">${files[i].name}</span><span class="badge badge-err">Error</span></div><p style="font-size:13px;color:#ff6b6b;margin-top:4px">${err.message}</p>`;
    }
  }

  document.getElementById('extractBtn').disabled = false;
  if (extractedUIDs.length > 0) {
    copyAllRow.style.display = 'flex';
    document.getElementById('copyAllLabel').textContent = `${extractedUIDs.length} UID${extractedUIDs.length > 1 ? 's' : ''} found`;
  }
}

function copyVal(id, btn) {
  const val = document.getElementById(id)?.textContent || '';
  navigator.clipboard.writeText(val).then(() => {
    btn.textContent = 'Copied!'; btn.classList.add('copied');
    setTimeout(() => { btn.textContent = 'Copy'; btn.classList.remove('copied'); }, 1500);
  });
}

function copyAllUIDs() {
  const text = extractedUIDs.map(e => e.uid).join('\\n');
  navigator.clipboard.writeText(text).then(() => {
    const btn = event.target;
    const orig = btn.innerHTML;
    btn.innerHTML = '✓ Copied!';
    setTimeout(() => { btn.innerHTML = orig; }, 1500);
  });
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
                "max_tokens": 500,
                "system": "You are a precise data extractor. Extract ONLY the UID/UserID number from TikTok settings screenshots. Return ONLY a JSON object with keys: uid (digits only string), did (DID digits only string or null), source (label shown like UID or UserId). No markdown, no extra text. If not found: {\"uid\":null,\"did\":null,\"source\":null}.",
                "messages": [{
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": image_data}},
                        {"type": "text", "text": "Extract the UID from this TikTok screenshot."}
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

    print(f"TikTok UID Extractor running at {url}")
    print("Close this window to stop the app.")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass

if __name__ == '__main__':
    run()
