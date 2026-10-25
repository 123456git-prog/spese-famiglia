/*
 * Lettore minimale di fogli di calcolo, senza librerie esterne.
 * Formati: .xlsx/.xlsm, CSV, tabelle HTML salvate come .xls, XML Spreadsheet 2003.
 * Restituisce le righe del primo foglio come array di array.
 * Le celle formattate come data diventano stringhe "AAAA-MM-GG".
 */
'use strict';

const XlsxLite = (() => {
  const serialToISO = n => new Date(Math.round((n - 25569) * 86400000)).toISOString().slice(0, 10);

  // ---------- ZIP ----------
  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === 'undefined') {
      throw new Error('Questo browser è troppo vecchio per leggere file .xlsx: aggiornalo oppure usa un file CSV.');
    }
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  function openZip(buf) {
    const u8 = new Uint8Array(buf);
    const dv = new DataView(buf);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('Il file non è un .xlsx valido.');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const files = new Map();
    const names = new TextDecoder();
    for (let n = 0; n < count && dv.getUint32(p, true) === 0x02014b50; n++) {
      const nameLen = dv.getUint16(p + 28, true);
      const name = names.decode(u8.subarray(p + 46, p + 46 + nameLen));
      files.set(name.toLowerCase(), {
        method: dv.getUint16(p + 10, true),
        size: dv.getUint32(p + 20, true),
        offset: dv.getUint32(p + 42, true),
      });
      p += 46 + nameLen + dv.getUint16(p + 30, true) + dv.getUint16(p + 32, true);
    }
    return {
      async text(name) {
        const f = files.get(name.toLowerCase());
        if (!f) return null;
        const start = f.offset + 30 + dv.getUint16(f.offset + 26, true) + dv.getUint16(f.offset + 28, true);
        const data = u8.subarray(start, start + f.size);
        let out;
        if (f.method === 0) out = data;
        else if (f.method === 8) out = await inflateRaw(data);
        else throw new Error('Compressione del file non supportata.');
        return new TextDecoder('utf-8').decode(out);
      },
    };
  }

  // ---------- XLSX ----------
  const xml = s => new DOMParser().parseFromString(s, 'application/xml');
  const tags = (node, name) => Array.from(node.getElementsByTagNameNS('*', name));
  const kids = (node, name) => Array.from(node.children).filter(c => c.localName === name);
  const textOf = node => tags(node, 't').filter(t => t.parentNode.localName !== 'rPh').map(t => t.textContent).join('');

  function colIndex(ref) {
    let n = 0;
    for (const ch of ref) {
      const c = ch.charCodeAt(0);
      if (c < 65 || c > 90) break;
      n = n * 26 + (c - 64);
    }
    return n - 1;
  }

  const BUILTIN_DATE_FORMATS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 45, 46, 47, 50, 57]);
  function isDateFormat(code) {
    const clean = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '');
    return /[dmy]/i.test(clean) && !/^general$/i.test(clean.trim());
  }

  async function readXlsx(buf) {
    const zip = openZip(buf);
    const wbText = await zip.text('xl/workbook.xml');
    if (!wbText) throw new Error('Il file non contiene un foglio Excel.');
    const wb = xml(wbText);
    const firstSheet = tags(wb, 'sheet')[0];
    let sheetPath = 'xl/worksheets/sheet1.xml';
    let sharedPath = 'xl/sharedStrings.xml';
    let stylesPath = 'xl/styles.xml';

    const relsText = await zip.text('xl/_rels/workbook.xml.rels');
    if (relsText) {
      const resolve = t => (t.startsWith('/') ? t.slice(1) : 'xl/' + t.replace(/^\.\//, ''));
      const rid = firstSheet && (firstSheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || firstSheet.getAttribute('r:id'));
      for (const rel of tags(xml(relsText), 'Relationship')) {
        const type = rel.getAttribute('Type') || '';
        const target = resolve(rel.getAttribute('Target') || '');
        if (rid && rel.getAttribute('Id') === rid) sheetPath = target;
        else if (type.endsWith('/sharedStrings')) sharedPath = target;
        else if (type.endsWith('/styles')) stylesPath = target;
      }
    }

    const shared = [];
    const ssText = await zip.text(sharedPath);
    if (ssText) for (const si of tags(xml(ssText), 'si')) shared.push(textOf(si));

    const dateStyle = [];
    const stText = await zip.text(stylesPath);
    if (stText) {
      const st = xml(stText);
      const custom = new Map(tags(st, 'numFmt').map(f => [+f.getAttribute('numFmtId'), f.getAttribute('formatCode') || '']));
      const xfsNode = tags(st, 'cellXfs')[0];
      if (xfsNode) {
        kids(xfsNode, 'xf').forEach((xf, i) => {
          const id = +xf.getAttribute('numFmtId') || 0;
          dateStyle[i] = BUILTIN_DATE_FORMATS.has(id) || (custom.has(id) && isDateFormat(custom.get(id)));
        });
      }
    }

    const sheetText = await zip.text(sheetPath);
    if (!sheetText) throw new Error('Impossibile leggere il primo foglio del file.');
    const rows = [];
    let lastRow = 0;
    for (const rowNode of tags(xml(sheetText), 'row')) {
      const r = +rowNode.getAttribute('r') || lastRow + 1;
      lastRow = r;
      const cells = [];
      let lastCol = -1;
      for (const c of kids(rowNode, 'c')) {
        const ref = c.getAttribute('r');
        const col = ref ? colIndex(ref) : lastCol + 1;
        lastCol = col;
        const t = c.getAttribute('t');
        const vNode = kids(c, 'v')[0];
        const v = vNode ? vNode.textContent : null;
        let val = '';
        if (t === 's') val = shared[+v] ?? '';
        else if (t === 'inlineStr') { const is = kids(c, 'is')[0]; val = is ? textOf(is) : ''; }
        else if (t === 'str' || t === 'd') val = v ?? '';
        else if (t === 'b') val = v === '1';
        else if (t === 'e') val = '';
        else if (v !== null && v !== '') {
          const num = Number(v);
          val = dateStyle[+c.getAttribute('s') || 0] && num > 0 ? serialToISO(num) : num;
        }
        cells[col] = val;
      }
      for (let i = 0; i < cells.length; i++) if (cells[i] === undefined) cells[i] = '';
      rows[r - 1] = cells;
    }
    for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
    return rows;
  }

  // ---------- CSV / HTML / XML 2003 ----------
  function decodeText(buf) {
    let text;
    try { text = new TextDecoder('utf-8', { fatal: true }).decode(buf); }
    catch { text = new TextDecoder('windows-1252').decode(buf); }
    return text.replace(/^﻿/, '');
  }

  function parseCSV(text) {
    const sample = text.split(/\r?\n/).filter(l => l.trim()).slice(0, 10).join('\n');
    const count = ch => sample.split(ch).length - 1;
    const delim = [';', '\t', ','].sort((a, b) => count(b) - count(a))[0];
    const rows = [];
    let row = [], field = '', quoted = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (quoted) {
        if (ch === '"' && text[i + 1] === '"') { field += '"'; i++; }
        else if (ch === '"') quoted = false;
        else field += ch;
      } else if (ch === '"' && field === '') quoted = true;
      else if (ch === delim) { row.push(field); field = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(field); rows.push(row); row = []; field = '';
      } else field += ch;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return rows.map(r => r.map(v => v.trim()));
  }

  function parseHTMLTable(text) {
    const doc = new DOMParser().parseFromString(text, 'text/html');
    return Array.from(doc.querySelectorAll('tr')).map(tr =>
      Array.from(tr.querySelectorAll('th,td')).map(td => td.textContent.replace(/\s+/g, ' ').trim()));
  }

  function parseXml2003(text) {
    const doc = xml(text);
    return tags(doc, 'Row').map(rowNode => {
      const out = [];
      let col = 0;
      for (const cell of kids(rowNode, 'Cell')) {
        const idx = cell.getAttribute('ss:Index') || cell.getAttributeNS('urn:schemas-microsoft-com:office:spreadsheet', 'Index');
        if (idx) col = +idx - 1;
        const data = kids(cell, 'Data')[0];
        let val = data ? data.textContent : '';
        const type = data && (data.getAttribute('ss:Type') || data.getAttributeNS('urn:schemas-microsoft-com:office:spreadsheet', 'Type'));
        if (type === 'Number') val = Number(val);
        else if (type === 'DateTime') val = val.slice(0, 10);
        out[col++] = val;
      }
      for (let i = 0; i < out.length; i++) if (out[i] === undefined) out[i] = '';
      return out;
    });
  }

  async function read(file) {
    const buf = await file.arrayBuffer();
    const head = new Uint8Array(buf.slice(0, 8));
    if (head[0] === 0x50 && head[1] === 0x4b) return readXlsx(buf);
    if (head[0] === 0xd0 && head[1] === 0xcf && head[2] === 0x11 && head[3] === 0xe0) {
      throw new Error('Il formato .xls (Excel 97-2003) non è supportato: dalla banca scarica il file in formato .xlsx oppure CSV, o aprilo con Excel e salvalo come .xlsx.');
    }
    const text = decodeText(buf);
    if (/<table[\s>]/i.test(text)) return parseHTMLTable(text);
    if (/urn:schemas-microsoft-com:office:spreadsheet/.test(text)) return parseXml2003(text);
    return parseCSV(text);
  }

  return { read, serialToISO };
})();
