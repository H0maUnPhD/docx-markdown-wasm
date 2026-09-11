const $ = (s) => document.querySelector(s);
const file = $('#file');
const drop = $('#drop');
const out = $('#out');
const preview = $('#preview');
const stats = $('#stats');
const engine = $('#engine');

let wasmConvert = null;
try {
  const m = await import('../pkg/docx_markdown_wasm.js');
  await m.default();
  wasmConvert = m.convert_docx;
  engine.textContent = 'Engine: Rust + WebAssembly';
  engine.classList.add('ok');
} catch (e) {
  engine.textContent = 'Engine: browser fallback (Rust WASM not built yet)';
  engine.classList.add('warn');
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
  if (eocd < 0) throw Error('Not a ZIP/DOCX file');
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
      throw Error('Unsupported ZIP compression method ' + method);
    }
    p += 46 + nlen + xlen + clen;
  }
  throw Error(wanted + ' not found');
}
function esc(s) {
  return s.replace(/\\/g, '\\\\').replace(/\*/g, '\\*').replace(/_/g, '\\_');
}
function jsConvert(xmlText) {
  const xml = new DOMParser().parseFromString(xmlText, 'application/xml');
  if (xml.querySelector('parsererror')) throw Error('Invalid Word XML');
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
  const ab = await f.arrayBuffer();
  let r;
  if (wasmConvert) r = wasmConvert(new Uint8Array(ab));
  else {
    const xml = await unzipEntry(ab, 'word/document.xml');
    r = jsConvert(new TextDecoder().decode(xml));
  }
  out.value = r.markdown;
  preview.textContent = r.markdown;
  stats.innerHTML = `<span class="stat">${f.name}</span><span class="stat">${(f.size / 1024).toFixed(1)} KB</span><span class="stat">${r.paragraphs} paragraphs</span><span class="stat">${r.headings} headings</span><span class="stat">${r.tables} tables</span><span class="stat">${r.words} words</span>`;
}

file.onchange = () => file.files[0] && convert(file.files[0]).catch(e => alert(e.message));
drop.onclick = () => file.click();
for (const ev of ['dragenter', 'dragover']) drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('drag'); });
for (const ev of ['dragleave', 'drop']) drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('drag'); });
drop.addEventListener('drop', e => {
  const f = e.dataTransfer.files[0];
  if (f) convert(f).catch(x => alert(x.message));
});
$('#copy').onclick = async () => { await navigator.clipboard.writeText(out.value); };
$('#save').onclick = () => {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([out.value], { type: 'text/markdown' }));
  a.download = 'converted.md';
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
};
