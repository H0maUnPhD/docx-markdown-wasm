const $ = (s) => document.querySelector(s);
const file = $('#file');
const drop = $('#drop');
const browse = $('#browse');
const changeFile = $('#changeFile');
const out = $('#out');
const preview = $('#preview');
const stats = $('#stats');
const engine = $('#engine');
const fileInfo = $('#fileInfo');
const fileName = $('#fileName');
const fileMeta = $('#fileMeta');
const copyBtn = $('#copy');
const saveBtn = $('#save');
const toast = $('#toast');
const toastText = $('#toastText');

let wasmConvert = null;
let currentFile = null;

function setEngine(type, label) {
  const klass = type === 'wasm'
    ? 'badge badge-success badge-outline h-9 gap-2 px-3 font-medium'
    : 'badge badge-warning badge-outline h-9 gap-2 px-3 font-medium';
  engine.className = klass;
  engine.innerHTML = type === 'wasm'
    ? '<span class="status status-success"></span><span>Rust + WebAssembly</span>'
    : '<span class="status status-warning"></span><span>Browser fallback</span>';
  engine.title = label;
}

try {
  const wasmUrl = new URL('./pkg/docx_markdown_wasm.js', import.meta.url);
  const m = await import(wasmUrl.href);
  await m.default();
  wasmConvert = m.convert_docx;
  setEngine('wasm', 'Rust/WASM engine loaded successfully');
} catch (e) {
  console.warn('Rust/WASM engine unavailable; using browser fallback.', e);
  setEngine('fallback', 'Rust/WASM could not be loaded; using browser fallback');
}

function showToast(message) {
  toastText.textContent = message;
  toast.classList.remove('hidden');
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => toast.classList.add('hidden'), 1800);
}

function setBusy(busy) {
  browse.disabled = busy;
  changeFile.disabled = busy;
  if (busy) {
    browse.innerHTML = '<span class="loading loading-spinner loading-xs"></span> Converting…';
  } else {
    browse.innerHTML = '<svg viewBox="0 0 24 24" class="size-4" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M3 7.5A2.5 2.5 0 0 1 5.5 5h3l2 2h8A2.5 2.5 0 0 1 21 9.5v7A2.5 2.5 0 0 1 18.5 19h-13A2.5 2.5 0 0 1 3 16.5z"/></svg> Choose DOCX';
  }
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function showFile(f) {
  currentFile = f;
  fileName.textContent = f.name;
  fileMeta.textContent = `${formatBytes(f.size)} • ${f.type || 'Word document'}`;
  fileInfo.classList.remove('hidden');
}

function renderStats(r) {
  const items = [
    ['Paragraphs', r.paragraphs],
    ['Headings', r.headings],
    ['Tables', r.tables],
    ['Words', r.words],
    ['Engine', wasmConvert ? 'WASM' : 'Browser']
  ];
  stats.innerHTML = items.map(([label, value]) => `
    <div class="surface rounded-2xl px-4 py-3">
      <div class="text-xs font-medium text-base-content/45">${label}</div>
      <div class="mt-1 text-lg font-semibold tracking-tight">${value}</div>
    </div>`).join('');
  stats.classList.remove('hidden');
  stats.classList.add('grid');
}

function u16(v, o) { return v.getUint16(o, true); }
function u32(v, o) { return v.getUint32(o, true); }
async function inflateRaw(data) {
  const ds = new DecompressionStream('deflate-raw');
  return new Uint8Array(await new Response(new Blob([data]).stream().pipeThrough(ds)).arrayBuffer());
}
async function unzipEntry(buffer, wanted) {
  const v = new DataView(buffer);
  const bytes = new Uint8Array(buffer);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (u32(v, i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw Error('This file is not a valid DOCX document.');
  const count = u16(v, eocd + 10);
  const cd = u32(v, eocd + 16);
  let p = cd;
  const dec = new TextDecoder();
  for (let i = 0; i < count; i++) {
    if (u32(v, p) !== 0x02014b50) break;
    const method = u16(v, p + 10);
    const csize = u32(v, p + 20);
    const nlen = u16(v, p + 28);
    const xlen = u16(v, p + 30);
    const clen = u16(v, p + 32);
    const loff = u32(v, p + 42);
    const name = dec.decode(bytes.slice(p + 46, p + 46 + nlen));
    if (name === wanted) {
      const ln = u16(v, loff + 26);
      const lx = u16(v, loff + 28);
      const start = loff + 30 + ln + lx;
      const comp = bytes.slice(start, start + csize);
      if (method === 0) return comp;
      if (method === 8) return inflateRaw(comp);
      throw Error(`Unsupported ZIP compression method ${method}`);
    }
    p += 46 + nlen + xlen + clen;
  }
  throw Error(`${wanted} not found in this DOCX file.`);
}
function esc(s) {
  return s.replace(/\\/g, '\\\\').replace(/\*/g, '\\*').replace(/_/g, '\\_');
}
function jsConvert(xmlText) {
  const xml = new DOMParser().parseFromString(xmlText, 'application/xml');
  if (xml.querySelector('parsererror')) throw Error('Invalid Word document XML.');
  const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const body = xml.getElementsByTagNameNS(NS, 'body')[0];
  let md = [], pars = 0, heads = 0, tables = 0;
  const textOf = n => Array.from(n.getElementsByTagNameNS(NS, 't')).map(x => x.textContent || '').join('');
  for (const n of Array.from(body.children)) {
    if (n.localName === 'p') {
      const t = textOf(n).trim();
      if (!t) continue;
      pars++;
      let st = '';
      const ps = n.getElementsByTagNameNS(NS, 'pStyle')[0];
      if (ps) st = ps.getAttributeNS(NS, 'val') || ps.getAttribute('w:val') || '';
      const l = st.toLowerCase();
      if (l.includes('heading1') || l === 'title') { md.push('# ' + esc(t)); heads++; }
      else if (l.includes('heading2')) { md.push('## ' + esc(t)); heads++; }
      else if (l.includes('heading3')) { md.push('### ' + esc(t)); heads++; }
      else md.push(esc(t));
    } else if (n.localName === 'tbl') {
      tables++;
      const rows = Array.from(n.getElementsByTagNameNS(NS, 'tr')).map(r =>
        Array.from(r.getElementsByTagNameNS(NS, 'tc')).map(c => textOf(c).trim().replace(/\|/g, '\\|'))
      );
      if (rows.length) {
        const cols = rows[0].length;
        md.push('| ' + rows[0].join(' | ') + ' |');
        md.push('| ' + Array(cols).fill('---').join(' | ') + ' |');
        for (const r0 of rows.slice(1)) {
          const r = [...r0];
          while (r.length < cols) r.push('');
          md.push('| ' + r.slice(0, cols).join(' | ') + ' |');
        }
      }
    }
  }
  const markdown = md.join('\n\n');
  return {
    markdown,
    paragraphs: pars,
    headings: heads,
    tables,
    words: markdown.trim() ? markdown.trim().split(/\s+/).length : 0
  };
}

async function convert(f) {
  if (!f.name.toLowerCase().endsWith('.docx')) throw Error('Please choose a .docx file.');
  showFile(f);
  setBusy(true);
  try {
    const ab = await f.arrayBuffer();
    let r;
    if (wasmConvert) r = wasmConvert(new Uint8Array(ab));
    else {
      const xml = await unzipEntry(ab, 'word/document.xml');
      r = jsConvert(new TextDecoder().decode(xml));
    }
    out.value = r.markdown;
    preview.textContent = r.markdown || 'The document did not contain convertible text.';
    renderStats(r);
    copyBtn.disabled = !r.markdown;
    saveBtn.disabled = !r.markdown;
    showToast('Document converted');
  } finally {
    setBusy(false);
  }
}

function openPicker() { file.click(); }
function handleError(error) {
  console.error(error);
  showToast(error?.message || 'Conversion failed');
}

browse.addEventListener('click', (e) => { e.stopPropagation(); openPicker(); });
changeFile.addEventListener('click', openPicker);
drop.addEventListener('click', openPicker);
drop.addEventListener('keydown', e => {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    openPicker();
  }
});
file.addEventListener('change', () => file.files[0] && convert(file.files[0]).catch(handleError));

for (const ev of ['dragenter', 'dragover']) {
  drop.addEventListener(ev, e => {
    e.preventDefault();
    e.stopPropagation();
    drop.classList.add('drag');
  });
}
for (const ev of ['dragleave', 'drop']) {
  drop.addEventListener(ev, e => {
    e.preventDefault();
    e.stopPropagation();
    drop.classList.remove('drag');
  });
}
drop.addEventListener('drop', e => {
  const f = e.dataTransfer?.files?.[0];
  if (f) convert(f).catch(handleError);
});

copyBtn.addEventListener('click', async () => {
  if (!out.value) return;
  await navigator.clipboard.writeText(out.value);
  const label = copyBtn.querySelector('span');
  const previous = label.textContent;
  label.textContent = 'Copied';
  showToast('Markdown copied');
  setTimeout(() => { label.textContent = previous; }, 1400);
});

saveBtn.addEventListener('click', () => {
  if (!out.value) return;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([out.value], { type: 'text/markdown;charset=utf-8' }));
  a.download = currentFile ? `${currentFile.name.replace(/\.docx$/i, '')}.md` : 'converted.md';
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  showToast('Markdown downloaded');
});

out.addEventListener('input', () => {
  preview.textContent = out.value || 'Choose a DOCX file to start.';
  copyBtn.disabled = !out.value;
  saveBtn.disabled = !out.value;
});
