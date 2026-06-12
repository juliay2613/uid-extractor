import threading
import webbrowser
import json
import urllib.request
import urllib.error
import socket
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
    --muted: #888; --accent: #ff2d55; --success: #00d68f;
    --mono: 'IBM Plex Mono', monospace; --sans: 'IBM Plex Sans', sans-serif;
  }
  body { font-family: var(--sans); background: var(--bg); color: var(--text); min-height: 100vh; display: flex; flex-direction: column; align-items: center; padding: 40px 20px 80px; }
  header { width: 100%; max-width: 640px; margin-bottom: 40px; display: flex; align-items: flex-end; gap: 16px; border-bottom: 1px solid var(--border); padding-bottom: 20px; }
  .logo { width: 36px; height: 36px; background: var(--accent); border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
  .logo svg { width: 20px; height: 20px; fill: white; }
  h1 { font-size: 18px; font-weight: 500; }
  h1 span { display: block; font-size: 12px; font-weight: 300; color: var(--muted); margin-top: 4px; letter-spacing: 0.05em; text-transform: uppercase; }
  .main { width: 100%; max-width: 640px; }
  .card { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 20px; margin-bottom: 16px; }
  .label { font-size: 11px; font-weight: 500; letter-spacing: 0.08em; text-transform: uppercase; color: var(--muted); margin-bottom: 10px; }
  .row { display: flex; gap: 8px; align-items: center; }
  input[type=password], input[type=text] { flex: 1; background: var(--surface2); border: 1px solid var(--border2); border-radius: 8px; padding: 10px 14px; font-family: var(--mono); font-size: 13px; color: var(--text); outline: none; }
  input:focus { border-color: var(--accent); }
  input::placeholder { color: #444; }
  .hint { font-size: 12px; color: var(--muted); margin-top: 8px; }
  .hint a { color: #ff6b6b; text-decoration: none; }
  .drop { background: var(--surface); border: 1px dashed var(--border2); border-radius: 12px; padding: 36px 20px; text-align: center; cursor: pointer; margin-bottom: 14px; transition: border-color 0.15s, background 0.15s; }
  .drop:hover, .drop.over { border-color: var(--accent); background: #1a1212; }
  .drop p { color: var(--muted); font-size: 14px; }
  .drop p strong { color: var(--text); }
  .drop small { font-size: 12px; color: #555; display: block; margin-top: 4px; }
  #thumbs { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 14px; }
  .tw { position: relative; width: 80px; }
  .tw img { width: 80px; height: 100px; object-fit: cover; border-radius: 8px; border: 1px solid var(--border2); display: block; }
  .tw span { font-size: 10px; color: var(--muted); display: block; margin-top: 3px; text-align: center; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .tw button { position: absolute; top: -5px; right: -5px; width: 18px; height: 18px; background: var(--accent); border: none; border-radius: 50%; color: white; font-size: 11px; cursor: pointer; }
  .btns { display: flex; gap: 10px; margin-bottom: 20px; flex-wrap: wrap; }
  .btn { background: var(--accent); color: white; border: none; border-radius: 8px; padding: 10px 20px; font-family: var(--sans); font-size: 14px; font-weight: 500; cursor: pointer; display: flex; align-items: center; gap: 6px; }
  .btn:hover { opacity: 0.88; }
  .btn:disabled { opacity: 0.4; cursor: not-allowed; }
  .btn.green { background: #1a8a5a; }
  .btn.ghost { background: transparent; color: var(--muted); border: 1px solid var(--border2); }
  .btn.ghost:hover { color: var(--text); }
  .rc { background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 14px 16px; margin-bottom: 8px; animation: fi 0.2s ease; }
  @keyframes fi { from { opacity: 0; transform: translateY(5px); } to { opacity: 1; transform: none; } }
  .rh { display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; }
  .rh span { font-size: 12px; color: var(--muted); }
  .badge { font-size: 11px; font-weight: 500; padding: 3px 8px; border-radius: 5px; }
  .ok { background: #0a2e20; color: var(--success); }
  .err { background: #2e0a10; color: #ff6b6b; }
  .loading { background: var(--surface2); color: var(--muted); }
  .ur { display: flex; align-items: center; justify-content: space-between; gap: 10px; padding: 10px 12px; background: var(--surface2); border-radius: 8px; margin-bottom: 5px; }
  .ul { font-size: 11px; color: var(--muted); margin-bottom: 2px; text-transform: uppercase; }
  .uv { font-family: var(--mono); font-size: 16px; font-weight: 500; }
  .cb { background: transparent; border: 1px solid var(--border2); color: var(--muted); border-radius: 6px; padding: 4px 10px; font-size: 12px; cursor: pointer; flex-shrink: 0; }
  .cb:hover { color: var(--text); }
  .cb.copied { color: var(--success); border-color: var(--success); }
  .spin { display: inline-block; width: 13px; height: 13px; border: 2px solid var(--border2); border-top-color: var(--muted); border-radius: 50%; animation: sp 0.7s linear infinite; vertical-align: -2px; margin-right: 5px; }
  @keyframes sp { to { transform: rotate(360deg); } }
  .bottom { margin-top: 14px; padding-top: 14px; border-top: 1px solid var(--border); display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px; }
  .bottom span { font-size: 13px; color: var(--muted); }
  footer { margin-top: 50px; font-size: 12px; color: #444; text-align: center; }
</style>
</head>
<body>
<header>
  <div class="logo"><svg viewBox="0 0 24 24"><path d="M19.59 6.69a4.83 4.83 0 0 1-3.77-4.25V2h-3.45v13.67a2.89 2.89 0 0 1-2.88 2.5 2.89 2.89 0 0 1-2.89-2.89 2.89 2.89 0 0 1 2.89-2.89c.28 0 .54.04.79.1V9.01a6.33 6.33 0 0 0-.79-.05 6.34 6.34 0 0 0-6.34 6.34 6.34 6.34 0 0 0 6.34 6.34 6.34 6.34 0 0 0 6.33-6.34l-.03-8.27a8.2 8.2 0 0 0 4.79 1.53V5.12a4.85 4.85 0 0 1-1-.43z"/></svg></div>
  <h1>UID Extractor <span>TikTok Settings Screenshot Tool</span></h1>
</header>
<div class="main">
  <div class="card">
    <div class="label">Anthropic API Key</div>
    <div class="row">
      <input type="password" id="apiKey" placeholder="sk-ant-api03-...">
      <button class="btn ghost" style="padding:8px 14px;font-size:13px" onclick="toggleKey(this)">Show</button>
      <button class="btn ghost" style="padding:8px 14px;font-size:13px" onclick="saveKey()">Save</button>
    </div>
    <p class="hint">Key saved locally on your device. Get one at <a href="https://console.anthropic.com/settings/keys" target="_blank">console.anthropic.com</a></p>
  </div>
  <div class="drop" id="drop" onclick="document.getElementById('fi').click()">
    <p><strong>Drop screenshots here</strong> or click to browse</p>
    <small>PNG, JPG — multiple files OK</small>
  </div>
  <input type="file" id="fi" accept="image/*" multiple style="display:none">
  <div id="thumbs"></div>
  <div class="btns" id="btns" style="display:none">
    <button class="btn" id="extBtn" onclick="extractAll()">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
      Extract UIDs
    </button>
    <button class="btn ghost" onclick="clearAll()">Clear all</button>
  </div>
  <div id="results"></div>
  <div class="bottom" id="bottom" style="display:none">
    <span id="foundLabel"></span>
    <div style="display:flex;gap:8px;flex-wrap:wrap">
      <button class="btn" style="padding:8px 16px;font-size:13px" onclick="copyAll(this)">Copy all UIDs</button>
      <button class="btn green" style="padding:8px 16px;font-size:13px" onclick="exportCSV()">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
        Export to Excel
      </button>
    </div>
  </div>
</div>
<footer>Running locally on your PC — API key never leaves your device.</footer>
<script>
const files = [];
let uids = [];

const saved = localStorage.getItem('tt_api_key');
if (saved) document.getElementById('apiKey').value = saved;

const drop = document.getElementById('drop');
const fi = document.getElementById('fi');
drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('over'); });
drop.addEventListener('dragleave', () => drop.classList.remove('over'));
drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('over'); addFiles(e.dataTransfer.files); });
fi.addEventListener('change', e => { addFiles(e.target.files); fi.value = ''; });

function addFiles(list) {
  Array.from(list).forEach(f => {
    if (!f.type.startsWith('image/')) return;
    const i = files.length; files.push(f);
    const url = URL.createObjectURL(f);
    const w = document.createElement('div');
    w.className = 'tw'; w.id = 'tw' + i;
    w.innerHTML = `<button onclick="rm(${i})">x</button><img src="${url}"><span>${f.name}</span>`;
    document.getElementById('thumbs').appendChild(w);
  });
  document.getElementById('btns').style.display = files.length ? 'flex' : 'none';
}

function rm(i) { files.splice(i, 1); rebuildThumbs(); document.getElementById('btns').style.display = files.length ? 'flex' : 'none'; }

function rebuildThumbs() {
  const t = document.getElementById('thumbs'); t.innerHTML = '';
  files.forEach((f, i) => {
    const url = URL.createObjectURL(f);
    const w = document.createElement('div'); w.className = 'tw'; w.id = 'tw' + i;
    w.innerHTML = `<button onclick="rm(${i})">x</button><img src="${url}"><span>${f.name}</span>`;
    t.appendChild(w);
  });
}

function clearAll() {
  files.length = 0; uids = [];
  document.getElementById('thumbs').innerHTML = '';
  document.getElementById('results').innerHTML = '';
  document.getElementById('btns').style.display = 'none';
  document.getElementById('bottom').style.display = 'none';
}

function toggleKey(btn) {
  const inp = document.getElementById('apiKey');
  inp.type = inp.type === 'password' ? 'text' : 'password';
  btn.textContent = inp.type === 'password' ? 'Show' : 'Hide';
}

function saveKey() {
  const k = document.getElementById('apiKey').value.trim();
  if (k) { localStorage.setItem('tt_api_key', k); alert('Key saved!'); }
}

async function toB64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(',')[1]);
    r.onerror = rej;
    r.readAsDataURL(file);
  });
}

async function extractAll() {
  const key = document.getElementById('apiKey').value.trim();
  if (!key) { alert('Enter your Anthropic API key first.'); return; }
  document.getElementById('results').innerHTML = '';
  document.getElementById('bottom').style.display = 'none';
  uids = [];
  const btn = document.getElementById('extBtn');
  btn.disabled = true; btn.textContent = 'Extracting...';

  for (let i = 0; i < files.length; i++) {
    const res = document.getElementById('results');
    const card = document.createElement('div');
    card.className = 'rc'; card.id = 'rc' + i;
    card.innerHTML = `<div class="rh"><span>${files[i].name}</span><span class="badge loading"><span class="spin"></span>Analyzing...</span></div>`;
    res.appendChild(card);

    try {
      const b64 = await toB64(files[i]);
      const resp = await fetch('/api/extract', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ apiKey: key, imageData: b64, mediaType: files[i].type || 'image/jpeg' })
      });
      const data = await resp.json();
      if (data.error) throw new Error(typeof data.error === 'object' ? data.error.message : data.error);
      const raw = (data.content?.find(c => c.type === 'text')?.text || '').replace(/```json|```/g,'').trim();
      const parsed = JSON.parse(raw);

      if (parsed?.uid) {
        uids.push({ file: files[i].name, uid: parsed.uid });
        card.innerHTML = `
          <div class="rh"><span>${files[i].name}</span><span class="badge ok">Found</span></div>
          <div class="ur"><div><div class="ul">${parsed.source||'UID'}</div><div class="uv" id="u${i}">${parsed.uid}</div></div><button class="cb" onclick="cp('u${i}',this)">Copy</button></div>
          ${parsed.did?`<div class="ur"><div><div class="ul">DID</div><div class="uv" id="d${i}">${parsed.did}</div></div><button class="cb" onclick="cp('d${i}',this)">Copy</button></div>`:''}`;
      } else {
        card.innerHTML = `<div class="rh"><span>${files[i].name}</span><span class="badge err">Not found</span></div><p style="font-size:13px;color:#ff6b6b;margin-top:4px">No UID detected.</p>`;
      }
    } catch(e) {
      card.innerHTML = `<div class="rh"><span>${files[i].name}</span><span class="badge err">Error</span></div><p style="font-size:13px;color:#ff6b6b;margin-top:4px">${e.message}</p>`;
    }
  }

  btn.disabled = false; btn.textContent = 'Extract UIDs';
  if (uids.length) {
    document.getElementById('bottom').style.display = 'flex';
    document.getElementById('foundLabel').textContent = uids.length + ' UID' + (uids.length > 1 ? 's' : '') + ' found';
  }
}

function cp(id, btn) {
  navigator.clipboard.writeText(document.getElementById(id)?.textContent || '').then(() => {
    btn.textContent = 'Copied!'; btn.classList.add('copied');
    setTimeout(() => { btn.textContent = 'Copy'; btn.classList.remove('copied'); }, 1500);
  });
}

function copyAll(btn) {
  navigator.clipboard.writeText(uids.map(e => e.uid).join('\\n')).then(() => {
    btn.textContent = 'Copied!';
    setTimeout(() => btn.textContent = 'Copy all UIDs', 1500);
  });
}

function exportCSV() {
  const rows = [['Filename', 'UID'], ...uids.map(e => [e.file, e.uid])];
  const csv = rows.map(r => r.map(c => '"' + String(c).replace(/"/g, '""') + '"').join(',')).join('\\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
  a.download = 'tiktok_uids.csv';
  a.click();
}
</script>
</body>
</html>"""

class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def do_GET(self):
        if self.path in ('/', '/index.html'):
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
                self.send_response(e.code)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(e.read())
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode())
        else:
            self.send_response(404)
            self.end_headers()


def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        return s.getsockname()[1]


def run():
    port = find_free_port()
    server = HTTPServer(('127.0.0.1', port), Handler)
    url = f'http://127.0.0.1:{port}'
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    print(f'TikTok UID Extractor running at {url}')
    print('Close this window to quit.')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == '__main__':
    run()
