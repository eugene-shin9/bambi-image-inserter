// XLSX: 원본 zip의 XML을 직접 읽어 격자 미리보기, 저장은 그림 관련 파트만 추가(나머지 원본 그대로)
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_DRAWING = NS_R + '/drawing';
const REL_IMAGE = NS_R + '/image';
const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml';
const EMU = 9525; // 1px(96dpi) = 9525 EMU
const MAX_PREVIEW_W = 6000, MAX_PREVIEW_H = 9000;

const parseXml = s => new DOMParser().parseFromString(s, 'application/xml');
const kids = (el, local) => el ? [...el.children].filter(c => c.localName === local) : [];
const kid = (el, local) => kids(el, local)[0] || null;
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function resolvePath(baseFile, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = baseFile.split('/'); parts.pop();
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop(); else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}
const relsPathOf = p => { const i = p.lastIndexOf('/'); return p.slice(0, i + 1) + '_rels/' + p.slice(i + 1) + '.rels'; };
function relTarget(fromFile, toFile) {
  const a = fromFile.split('/'); a.pop();
  const b = toFile.split('/');
  let i = 0; while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
  return '../'.repeat(a.length - i) + b.slice(i).join('/');
}
async function readText(zip, path) {
  const f = zip.file(path) || zip.file(Object.keys(zip.files).find(k => k.toLowerCase() === path.toLowerCase()) || '\0');
  return f ? f.async('string') : null;
}
async function readRels(zip, path) {
  const s = await readText(zip, relsPathOf(path));
  if (!s) return [];
  return [...parseXml(s).getElementsByTagNameNS(NS_PKG_REL, 'Relationship')].map(r => ({
    id: r.getAttribute('Id'), type: r.getAttribute('Type'), target: r.getAttribute('Target'), mode: r.getAttribute('TargetMode')
  }));
}

// 글꼴별 최대 숫자 폭(MDW, px)
function mdwFor(name, sz) {
  const n = (name || '').toLowerCase().replace(/\s/g, '');
  const s = Math.round(sz || 11);
  const table = {
    'calibri': { 10: 7, 11: 7, 12: 8 }, 'arial': { 9: 6, 10: 7, 11: 8 },
    '맑은고딕': { 9: 7, 10: 7, 11: 8, 12: 9 }, 'malgungothic': { 9: 7, 10: 7, 11: 8, 12: 9 },
    '돋움': { 9: 7, 10: 7, 11: 8 }, 'dotum': { 9: 7, 10: 7, 11: 8 },
    '굴림': { 9: 7, 10: 7, 11: 8 }, 'gulim': { 9: 7, 10: 7, 11: 8 },
    '바탕': { 10: 7, 11: 8 }, 'batang': { 10: 7, 11: 8 }
  };
  if (table[n] && table[n][s]) return table[n][s];
  try {
    const c = document.createElement('canvas').getContext('2d');
    c.font = `${s * 96 / 72}px "${name}", sans-serif`;
    let m = 0; for (const d of '0123456789') m = Math.max(m, c.measureText(d).width);
    if (m > 3 && m < 30) return Math.round(m);
  } catch (e) { }
  return 7;
}
function defaultRowPt(name, sz) {
  const n = (name || '').toLowerCase().replace(/\s/g, '');
  const s = sz || 11;
  if (/맑은고딕|malgun/.test(n)) return s <= 10 ? 15 : s <= 11 ? 16.5 : Math.round(s * 1.5);
  if (/calibri/.test(n)) return s <= 11 ? 15 : Math.round(s * 1.36);
  return Math.max(12.75, Math.round(s * 1.4 * 4) / 4);
}
const widthToPx = (w, mdw) => Math.trunc(((256 * w + Math.trunc(128 / mdw)) / 256) * mdw);
const colName = n => { let s = ''; n++; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const parseRef = ref => { const m = /^([A-Z]+)(\d+)$/.exec(ref || ''); if (!m) return null; let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return { c: c - 1, r: +m[2] - 1 }; };

export async function openXlsx(file) {
  const buf = await file.arrayBuffer();
  const head = new Uint8Array(buf.slice(0, 8));
  if (head[0] === 0xD0 && head[1] === 0xCF) throw new Error('암호가 걸렸거나 옛 형식(.xls)의 엑셀 파일입니다. 엑셀에서 암호를 풀고 "Excel 통합 문서(.xlsx)"로 저장한 뒤 올려 주세요.');
  let zip;
  const origDirs = new Set();
  try { zip = await window.JSZip.loadAsync(buf); Object.keys(zip.files).forEach(n => zip.files[n].dir && origDirs.add(n)); } catch (e) { throw new Error('엑셀 파일을 읽을 수 없습니다(손상되었거나 xlsx 형식이 아님).'); }
  const wbPath = 'xl/workbook.xml';
  const wbXml = await readText(zip, wbPath);
  if (!wbXml) throw new Error('엑셀 통합 문서 구조를 찾을 수 없습니다.');
  if (wbXml.includes('purl.oclc.org/ooxml')) throw new Error('"Strict Open XML" 형식은 지원하지 않습니다. 엑셀에서 일반 "Excel 통합 문서(.xlsx)"로 다시 저장해 주세요.');
  const wb = parseXml(wbXml);
  const wbRels = await readRels(zip, wbPath);
  const sheets = [];
  for (const s of wb.getElementsByTagName('*')) {
    if (s.localName !== 'sheet') continue;
    const rid = s.getAttributeNS(NS_R, 'id') || [...s.attributes].find(a => a.localName === 'id' && a.prefix)?.value;
    const rel = wbRels.find(r => r.id === rid);
    if (!rel || !/\/worksheet$/.test(rel.type)) continue; // 차트 시트 등 제외
    sheets.push({ name: s.getAttribute('name'), path: resolvePath(wbPath, rel.target), hidden: (s.getAttribute('state') || 'visible') !== 'visible' });
  }
  if (!sheets.length) throw new Error('그림을 넣을 수 있는 워크시트가 없습니다.');

  // 공유 문자열
  const sst = [];
  const sstRel = wbRels.find(r => /\/sharedStrings$/.test(r.type));
  if (sstRel) {
    const x = await readText(zip, resolvePath(wbPath, sstRel.target));
    if (x) for (const si of parseXml(x).documentElement.children) {
      if (si.localName !== 'si') continue;
      let t = '';
      const walk = el => { for (const c of el.children) { if (c.localName === 'rPh' || c.localName === 'phoneticPr') continue; if (c.localName === 't') t += c.textContent; else walk(c); } };
      walk(si); sst.push(t);
    }
  }
  // 스타일: 기본 글꼴 + 날짜 서식
  let fontName = 'Calibri', fontSz = 11; const dateXf = new Set();
  const stRel = wbRels.find(r => /\/styles$/.test(r.type));
  if (stRel) {
    const x = await readText(zip, resolvePath(wbPath, stRel.target));
    if (x) {
      const st = parseXml(x).documentElement;
      const f0 = kid(kid(st, 'fonts'), 'font');
      if (f0) { fontName = kid(f0, 'name')?.getAttribute('val') || fontName; fontSz = +(kid(f0, 'sz')?.getAttribute('val') || fontSz); }
      const custom = new Map(kids(kid(st, 'numFmts'), 'numFmt').map(n => [+n.getAttribute('numFmtId'), n.getAttribute('formatCode') || '']));
      kids(kid(st, 'cellXfs'), 'xf').forEach((xf, i) => {
        const id = +(xf.getAttribute('numFmtId') || 0);
        const code = custom.get(id);
        if ((id >= 14 && id <= 22) || (id >= 45 && id <= 47) || (code && /[ymd]/i.test(code.replace(/\[[^\]]*\]|"[^"]*"/g, '')) && !/^[#0.,%\s]*$/.test(code))) dateXf.add(i);
      });
    }
  }
  const MDW = mdwFor(fontName, fontSz);

  const loaded = new Map();
  async function loadSheet(i) {
    if (loaded.has(i)) return loaded.get(i);
    const sh = sheets[i];
    const xml = await readText(zip, sh.path);
    const doc = parseXml(xml).documentElement;
    const fmt = kid(doc, 'sheetFormatPr');
    const baseCW = +(fmt?.getAttribute('baseColWidth') || 8);
    const defCW = fmt?.getAttribute('defaultColWidth');
    const defColPx = defCW ? widthToPx(+defCW, MDW) : Math.trunc(baseCW * MDW + 5 + (baseCW === 8 ? 0.43 * MDW : 0));
    const defRowPx = Math.round((+(fmt?.getAttribute('defaultRowHeight') || 0) || defaultRowPt(fontName, fontSz)) * 96 / 72);
    const colW = new Map();
    for (const c of kids(kid(doc, 'cols'), 'col')) {
      const min = +c.getAttribute('min') - 1, max = Math.min(+c.getAttribute('max') - 1, 16383);
      const hid = /^(1|true)$/.test(c.getAttribute('hidden') || '');
      const w = c.getAttribute('width');
      const px = hid ? 0 : (w != null ? widthToPx(+w, MDW) : defColPx);
      for (let k = min; k <= max && k < 2000; k++) colW.set(k, px);
    }
    const rowH = new Map(); const cells = [];
    let maxR = 0, maxC = 0;
    for (const row of kids(kid(doc, 'sheetData'), 'row')) {
      const r = +row.getAttribute('r') - 1;
      const hid = /^(1|true)$/.test(row.getAttribute('hidden') || '');
      const ht = row.getAttribute('ht');
      if (hid) rowH.set(r, 0); else if (ht != null) rowH.set(r, Math.round(+ht * 96 / 72));
      for (const c of kids(row, 'c')) {
        const p = parseRef(c.getAttribute('r')); if (!p) continue;
        const t = c.getAttribute('t'); const v = kid(c, 'v')?.textContent;
        let text = '', num = false;
        if (t === 's') text = sst[+v] ?? '';
        else if (t === 'inlineStr') text = kid(c, 'is')?.textContent || '';
        else if (t === 'str' || t === 'e') text = v || '';
        else if (t === 'b') text = v === '1' ? 'TRUE' : 'FALSE';
        else if (v != null && v !== '') {
          num = true; const n = +v;
          if (dateXf.has(+(c.getAttribute('s') || 0)) && n > 0 && n < 2958466) {
            const d = new Date(Math.round((n - 25569) * 86400000));
            text = d.toISOString().slice(0, 10);
          } else text = Number.isInteger(n) ? String(n) : String(+n.toPrecision(10));
        }
        if (text !== '') { cells.push({ r: p.r, c: p.c, text, num }); maxR = Math.max(maxR, p.r); maxC = Math.max(maxC, p.c); }
      }
    }
    const merges = kids(kid(doc, 'mergeCells'), 'mergeCell').map(m => {
      const [a, b] = (m.getAttribute('ref') || '').split(':').map(parseRef);
      return a && b ? { r0: a.r, c0: a.c, r1: b.r, c1: b.c } : null;
    }).filter(Boolean);
    for (const m of merges) { maxR = Math.max(maxR, m.r1); maxC = Math.max(maxC, m.c1); }
    const cw = k => colW.has(k) ? colW.get(k) : defColPx;
    const rh = k => rowH.has(k) ? rowH.get(k) : defRowPx;
    // 미리보기 범위: 사용 범위 + 여유, 최대 크기 제한
    const colX = [0]; let x = 0, c = 0;
    while ((c <= maxC + 6 || x < 1200) && x < MAX_PREVIEW_W && c < 2000) { x += cw(c); colX.push(x); c++; }
    const rowY = [0]; let y = 0, r = 0;
    while ((r <= maxR + 15 || y < 900) && y < MAX_PREVIEW_H && r < 20000) { y += rh(r); rowY.push(y); r++; }
    const info = { xml, colX, rowY, cells, merges, truncated: (maxC >= colX.length - 1 || maxR >= rowY.length - 1) };
    loaded.set(i, info);
    return info;
  }

  const first = sheets.findIndex(s => !s.hidden);
  const api = {
    kind: 'xlsx', ext: /\.xlsm$/i.test(file.name) ? 'xlsm' : 'xlsx',
    mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    warnings: [], pxPerMm: 96 / 25.4, perSheet: true,
    sheets: sheets.map(s => ({ name: s.name, hidden: s.hidden })),
    startSheet: first < 0 ? 0 : first,
    pages: [],
    async loadSheetPage(i) {
      const s = await loadSheet(i);
      const w = s.colX[s.colX.length - 1], h = s.rowY[s.rowY.length - 1];
      if (s.truncated) api.note = '큰 시트라 미리보기는 앞부분만 보여 줍니다(그림은 보이는 범위에 넣을 수 있음).'; else api.note = '';
      return { w, h };
    },
    async renderPage(i, scale) {
      const s = await loadSheet(i);
      const W = s.colX[s.colX.length - 1], H = s.rowY[s.rowY.length - 1];
      const k = Math.min(scale, 16000 / W, 16000 / H, Math.sqrt(60e6 / (W * H)));
      const cv = document.createElement('canvas');
      cv.width = Math.ceil(W * k); cv.height = Math.ceil(H * k);
      const g = cv.getContext('2d');
      g.scale(k, k);
      g.fillStyle = '#fff'; g.fillRect(0, 0, W, H);
      // 병합 영역 안쪽 선은 그리지 않음
      const inMerge = (r, c) => s.merges.find(m => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1);
      g.strokeStyle = '#d9dde3'; g.lineWidth = 1 / k;
      g.beginPath();
      for (let ci = 0; ci < s.colX.length; ci++) {
        const X = s.colX[ci];
        for (let ri = 0; ri < s.rowY.length - 1; ri++) {
          const m = ci > 0 && ci < s.colX.length - 1 ? inMerge(ri, ci) : null;
          if (m && ci > m.c0) continue;
          g.moveTo(X, s.rowY[ri]); g.lineTo(X, s.rowY[ri + 1]);
        }
      }
      for (let ri = 0; ri < s.rowY.length; ri++) {
        const Y = s.rowY[ri];
        for (let ci = 0; ci < s.colX.length - 1; ci++) {
          const m = ri > 0 && ri < s.rowY.length - 1 ? inMerge(ri, ci) : null;
          if (m && ri > m.r0) continue;
          g.moveTo(s.colX[ci], Y); g.lineTo(s.colX[ci + 1], Y);
        }
      }
      g.stroke();
      g.fillStyle = '#1f2328'; g.textBaseline = 'middle';
      g.font = `${Math.round(fontSz * 96 / 72)}px "${fontName}", "Malgun Gothic", "Apple SD Gothic Neo", sans-serif`;
      for (const cell of s.cells) {
        if (cell.r >= s.rowY.length - 1 || cell.c >= s.colX.length - 1) continue;
        const m = s.merges.find(q => q.r0 === cell.r && q.c0 === cell.c);
        const x0 = s.colX[cell.c], y0 = s.rowY[cell.r];
        const x1 = s.colX[Math.min((m ? m.c1 : cell.c) + 1, s.colX.length - 1)];
        const y1 = s.rowY[Math.min((m ? m.r1 : cell.r) + 1, s.rowY.length - 1)];
        if (x1 - x0 < 2 || y1 - y0 < 2) continue;
        g.save(); g.beginPath(); g.rect(x0 + 1, y0, x1 - x0 - 2, y1 - y0); g.clip();
        if (cell.num) { g.textAlign = 'right'; g.fillText(cell.text, x1 - 3, (y0 + y1) / 2); }
        else { g.textAlign = m ? 'center' : 'left'; g.fillText(cell.text, m ? (x0 + x1) / 2 : x0 + 3, (y0 + y1) / 2); }
        g.restore();
      }
      return cv;
    },
    async snaps(i) { const s = await loadSheet(i); return { xs: s.colX.slice(), ys: s.rowY.slice() }; },
    cellLabel(i, x, y) {
      const s = loaded.get(i); if (!s) return '';
      const c = Math.max(0, upper(s.colX, x) - 1), r = Math.max(0, upper(s.rowY, y) - 1);
      return colName(c) + (r + 1);
    },
    async save(items, getImage) {
      const bySheet = new Map();
      for (const it of items) { if (!bySheet.has(it.page)) bySheet.set(it.page, []); bySheet.get(it.page).push(it); }
      let ctXml = await readText(zip, '[Content_Types].xml');
      const names = Object.keys(zip.files);
      const maxNum = re => names.reduce((m, n) => { const q = re.exec(n); return q ? Math.max(m, +q[1]) : m; }, 0);
      let imgNo = maxNum(/^xl\/media\/image(\d+)\.\w+$/i);
      let drawNo = maxNum(/^xl\/drawings\/drawing(\d+)\.xml$/i);
      const mediaByKey = new Map();
      const ensureDefault = (ext, type) => {
        if (new RegExp(`<Default\\b[^>]*Extension="${ext}"`, 'i').test(ctXml)) return;
        ctXml = ctXml.replace(/<Types\b[^>]*>/, m => `${m}<Default Extension="${ext}" ContentType="${type}"/>`);
      };
      for (const [si, list] of bySheet) {
        const s = await loadSheet(si);
        const sheetPath = sheets[si].path;
        let sheetXml = s.xml;
        const sheetRels = await readRels(zip, sheetPath);
        // 기존 drawing 연결 확인
        const dm = /<(?:\w+:)?drawing\b[^>]*?\b(\w+):id="([^"]+)"/.exec(sheetXml);
        let drawingPath, drawingXml, drawingRels;
        if (dm) {
          const rel = sheetRels.find(r => r.id === dm[2]);
          drawingPath = rel ? resolvePath(sheetPath, rel.target) : null;
        }
        if (drawingPath && zip.file(drawingPath)) {
          drawingXml = await readText(zip, drawingPath);
          drawingRels = await readRels(zip, drawingPath);
        } else {
          drawNo++;
          drawingPath = `xl/drawings/drawing${drawNo}.xml`;
          drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}"></xdr:wsDr>`;
          drawingRels = [];
          ctXml = ctXml.replace(/<\/Types>\s*$/, `<Override PartName="/${drawingPath}" ContentType="${CT_DRAWING}"/></Types>`);
          // 시트 rels에 drawing 연결 추가
          const rid = nextRid(sheetRels);
          sheetRels.push({ id: rid, type: REL_DRAWING, target: relTarget(sheetPath, drawingPath) });
          zip.file(relsPathOf(sheetPath), relsXml(sheetRels), { createFolders: false });
          sheetXml = insertDrawingEl(sheetXml, rid);
        }
        let maxId = 0;
        for (const m of drawingXml.matchAll(/<(?:\w+:)?cNvPr\b[^>]*\bid="(\d+)"/g)) maxId = Math.max(maxId, +m[1]);
        let anchors = '';
        for (const it of list) {
          const img = await getImage(it.imgId);
          let media = mediaByKey.get(img.key);
          if (!media) {
            imgNo++;
            const ext = img.ext === 'jpg' ? 'jpeg' : 'png';
            media = `xl/media/image${imgNo}.${ext}`;
            zip.file(media, img.bytes, { createFolders: false });
            ensureDefault(ext, ext === 'png' ? 'image/png' : 'image/jpeg');
            mediaByKey.set(img.key, media);
          }
          let rel = drawingRels.find(r => r.type === REL_IMAGE && resolvePath(drawingPath, r.target) === media);
          if (!rel) { rel = { id: nextRid(drawingRels), type: REL_IMAGE, target: relTarget(drawingPath, media) }; drawingRels.push(rel); }
          const x = Math.max(0, it.x), y = Math.max(0, it.y);
          const c = Math.max(0, upper(s.colX, x) - 1), r = Math.max(0, upper(s.rowY, y) - 1);
          const colOff = Math.round((x - s.colX[c]) * EMU), rowOff = Math.round((y - s.rowY[r]) * EMU);
          const cx = Math.round(it.w * EMU), cy = Math.round(it.h * EMU);
          const id = ++maxId;
          const nsAttr = dm ? ` xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}"` : '';
          anchors += `<xdr:oneCellAnchor${nsAttr}><xdr:from><xdr:col>${c}</xdr:col><xdr:colOff>${colOff}</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>${rowOff}</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/>` +
            `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${esc('그림 ' + id)}" descr="${esc(img.name || '')}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
            `<xdr:blipFill><a:blip xmlns:r="${NS_R}" r:embed="${rel.id}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
            `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
        }
        drawingXml = drawingXml.replace(/<\/((?:\w+:)?wsDr)>\s*$/, `${anchors}</$1>`);
        zip.file(drawingPath, drawingXml, { createFolders: false });
        zip.file(relsPathOf(drawingPath), relsXml(drawingRels), { createFolders: false });
        zip.file(sheetPath, sheetXml, { createFolders: false });
      }
      zip.file('[Content_Types].xml', ctXml, { createFolders: false });
      for (const n of Object.keys(zip.files)) if (zip.files[n].dir && !origDirs.has(n)) delete zip.files[n];
      const out = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
      // 저장 뒤 원본 zip 상태로 되돌리기(다시 저장해도 중복되지 않게)
      zip = await window.JSZip.loadAsync(buf);
      return new Blob([out], { type: api.mime });
    }
  };
  return api;
}

function upper(arr, v) { let lo = 0, hi = arr.length; while (lo < hi) { const m = (lo + hi) >> 1; if (arr[m] <= v) lo = m + 1; else hi = m; } return lo; }
function nextRid(rels) { let n = 0; for (const r of rels) { const m = /^rId(\d+)$/.exec(r.id); if (m) n = Math.max(n, +m[1]); } let id; do { id = 'rId' + (++n); } while (rels.some(r => r.id === id)); return id; }
function relsXml(rels) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG_REL}">` +
    rels.map(r => `<Relationship Id="${esc(r.id)}" Type="${esc(r.type)}" Target="${esc(r.target)}"${r.mode ? ` TargetMode="${esc(r.mode)}"` : ''}/>`).join('') + '</Relationships>';
}
// <drawing>은 스키마 순서상 legacyDrawing·picture·oleObjects·controls·tableParts·extLst 앞
function insertDrawingEl(xml, rid) {
  const rootM = /<((\w+):)?worksheet\b[^>]*>/.exec(xml);
  const pfx = rootM && rootM[2] ? rootM[2] + ':' : '';
  let rp = null;
  const decl = new RegExp(`xmlns:(\\w+)="${NS_R.replace(/[./]/g, '\\$&')}"`).exec(rootM ? rootM[0] : '');
  if (decl) rp = decl[1];
  if (!rp) { rp = 'r'; let n = 0; while (new RegExp(`xmlns:${rp}=`).test(rootM[0])) rp = 'r' + (++n); xml = xml.replace(rootM[0], rootM[0].replace(/>$/, ` xmlns:${rp}="${NS_R}">`).replace(/\/>$/, '/>')); }
  const el = `<${pfx}drawing ${rp}:id="${rid}"/>`;
  const afterData = Math.max(xml.search(/<\/(?:\w+:)?sheetData>/), xml.search(/<(?:\w+:)?sheetData\b[^>]*\/>/));
  const re = /<(?:\w+:)?(legacyDrawing|legacyDrawingHF|drawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b|<mc:AlternateContent\b|<\/(?:\w+:)?worksheet>/g;
  re.lastIndex = Math.max(0, afterData);
  const m = re.exec(xml);
  if (!m) throw new Error('엑셀 시트 구조를 해석하지 못했습니다.');
  return xml.slice(0, m.index) + el + xml.slice(m.index);
}
