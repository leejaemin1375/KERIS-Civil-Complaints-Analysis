// Minimal .docx writer (stored ZIP, no deps)
const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
export function zip(files) {
  const enc = new TextEncoder(), parts = [], cen = []; let off = 0;
  for (const [name, data] of files) {
    const nb = enc.encode(name), db = typeof data === 'string' ? enc.encode(data) : data, crc = crc32(db);
    const h = new DataView(new ArrayBuffer(30));
    h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint32(14, crc, true); h.setUint32(18, db.length, true); h.setUint32(22, db.length, true); h.setUint16(26, nb.length, true);
    parts.push(new Uint8Array(h.buffer), nb, db);
    const c = new DataView(new ArrayBuffer(46));
    c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint32(16, crc, true); c.setUint32(20, db.length, true); c.setUint32(24, db.length, true); c.setUint16(28, nb.length, true); c.setUint32(42, off, true);
    cen.push(new Uint8Array(c.buffer), nb);
    off += 30 + nb.length + db.length;
  }
  const cs = cen.reduce((a, b) => a + b.length, 0), e = new DataView(new ArrayBuffer(22));
  e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, cs, true); e.setUint32(16, off, true);
  return new Blob([...parts, ...cen, new Uint8Array(e.buffer)], { type: 'application/zip' });
}
const x = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const run = (t, o = {}) => `<w:r><w:rPr><w:rFonts w:ascii="Malgun Gothic" w:hAnsi="Malgun Gothic" w:eastAsia="맑은 고딕"/>${o.b ? '<w:b/>' : ''}${o.color ? `<w:color w:val="${o.color}"/>` : ''}<w:sz w:val="${o.sz || 21}"/></w:rPr><w:t xml:space="preserve">${x(t)}</w:t></w:r>`;
const para = (t, o = {}) => `<w:p><w:pPr><w:spacing w:before="${o.before || 0}" w:after="${o.after ?? 100}"/>${o.ind ? `<w:ind w:left="${o.ind}"/>` : ''}${o.center ? '<w:jc w:val="center"/>' : ''}${o.border ? '<w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" w:color="063A74"/></w:pBdr>' : ''}</w:pPr>${run(t, o)}</w:p>`;
const cell = (t, w, head) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${head ? '<w:shd w:val="clear" w:color="auto" w:fill="EEF2F7"/>' : ''}</w:tcPr>${para(t, { b: head, sz: 19, after: 0 })}</w:tc>`;
function table(rows, widths) {
  const b = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(s => `<w:${s} w:val="single" w:sz="4" w:color="B1B8BE"/>`).join('');
  return `<w:tbl><w:tblPr><w:tblW w:w="${widths.reduce((a, c) => a + c, 0)}" w:type="dxa"/><w:tblBorders>${b}</w:tblBorders><w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${widths.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows.map((r, i) => `<w:tr>${r.map((c, j) => cell(c, widths[j], i === 0)).join('')}</w:tr>`).join('')}</w:tbl>${para('', { after: 120 })}`;
}
// blocks: {t:'title'|'meta'|'h'|'sub'|'p'|'li'|'table', text, rows, widths}
export function docx(blocks) {
  const body = blocks.map(b => {
    if (b.t === 'title') return para(b.text, { b: true, sz: 34, color: '063A74', after: 120, border: true });
    if (b.t === 'meta') return para(b.text, { sz: 18, color: '464C53', after: 240 });
    if (b.t === 'h') return para(b.text, { b: true, sz: 26, color: '063A74', before: 240, after: 120 });
    if (b.t === 'sub') return para(b.text, { b: true, sz: 21, before: 80, after: 60 });
    if (b.t === 'cap') return para(b.text, { sz: 18, color: '464C53', center: true, after: 160 });
    if (b.t === 'li') return para(`- ${b.text}`, { ind: 360, after: 60 });
    if (b.t === 'table') return table(b.rows, b.widths || b.rows[0].map(() => Math.floor(9000 / b.rows[0].length)));
    return para(b.text);
  }).join('');
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  return zip([
    ['[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'],
    ['_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'],
    ['word/document.xml', doc]
  ]);
}
export async function blobBytes(b) { return new Uint8Array(await b.arrayBuffer()); }
export function save(blob, name, type) {
  const b = type ? new Blob([blob], { type }) : blob, a = document.createElement('a');
  a.href = URL.createObjectURL(b); a.download = name; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}
