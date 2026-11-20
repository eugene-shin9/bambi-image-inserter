import { loadImageFile, estimateBg, pickColor, solidCutout, aiCutout, encodeForInsert } from './cutout.js';

const $ = id => document.getElementById(id);
const S = {
  A: null, fileName: '', view: [], items: [], images: new Map(), sel: null,
  zoom: 1, fitMode: true, undo: [], redo: [], dirty: false, sheet: 0, snapCache: new Map(), seq: 1
};
const dpr = () => Math.min(3, window.devicePixelRatio || 1);
const uid = p => p + (S.seq++);

/* ---------- 알림 ---------- */
let toastT;
function toast(msg, ms = 2600) {
  const t = $('toast'); t.textContent = msg; t.classList.add('on');
  clearTimeout(toastT); if (ms) toastT = setTimeout(() => t.classList.remove('on'), ms);
}
function modal(text, buttons = [{ label: '확인', value: true, primary: true }]) {
  return new Promise(res => {
    $('mtext').textContent = text;
    const box = $('mbtns'); box.innerHTML = '';
    for (const b of buttons) {
      const el = document.createElement('button');
      el.className = 'btn' + (b.primary ? ' primary' : '');
      el.textContent = b.label;
      el.onclick = () => { $('modal').hidden = true; res(b.value); };
      box.appendChild(el);
    }
    $('modal').hidden = false; box.lastChild.focus();
  });
}
const confirmBox = t => modal(t, [{ label: '취소', value: false }, { label: '계속', value: true, primary: true }]);

/* ---------- 문서 열기 ---------- */
async function openDoc(file) {
  if (!file) return;
  if (S.items.length && S.dirty && !(await confirmBox('저장하지 않은 작업이 있습니다. 버리고 새 문서를 열까요?'))) return;
  const name = file.name || '';
  const ext = (/\.([^.]+)$/.exec(name.toLowerCase()) || [])[1];
  try {
    toast('문서를 여는 중…', 0);
    let A;
    if (ext === 'pdf') A = await (await import('./fmt-pdf.js')).openPdf(file);
    else if (ext === 'xlsx' || ext === 'xlsm') A = await (await import('./fmt-xlsx.js')).openXlsx(file);
    else if (ext === 'hwp' || ext === 'hwpx') A = await (await import('./fmt-hwp.js')).openHwp(file, m => toast(m, 0));
    else if (ext === 'xls') throw new Error('옛 엑셀 형식(.xls)은 지원하지 않습니다. 엑셀에서 "Excel 통합 문서(.xlsx)"로 저장한 뒤 올려 주세요.');
    else throw new Error('PDF, XLSX, HWP, HWPX 파일만 열 수 있습니다.');
    S.A = A; S.fileName = name; S.items = []; S.sel = null; S.undo = []; S.redo = []; S.dirty = false; S.snapCache.clear();
    $('fname').textContent = name;
    $('empty').hidden = true; $('saveBtn').disabled = false;
    if (A.perSheet) { buildTabs(); await showSheet(A.startSheet); }
    else { $('sheetTabs').hidden = true; S.view = A.pages.map((p, i) => ({ idx: i, w: p.w, h: p.h })); fitZoom(); buildPages(); }
    showNotice(A.warnings);
    toast(`${name} 열림`);
    updateUndoBtns(); updatePanel();
  } catch (e) {
    console.error(e); $('toast').classList.remove('on');
    await modal(e.message || String(e));
  }
}
function showNotice(list) {
  const n = $('notice'); const all = [...(list || [])]; if (S.A && S.A.note) all.push(S.A.note);
  n.hidden = !all.length; n.textContent = all.join(' ');
}
function buildTabs() {
  const t = $('sheetTabs'); t.innerHTML = ''; t.hidden = false;
  S.A.sheets.forEach((s, i) => {
    if (s.hidden) return;
    const b = document.createElement('button'); b.textContent = s.name; b.dataset.i = i;
    b.onclick = () => showSheet(i); t.appendChild(b);
  });
}
async function showSheet(i) {
  S.sheet = i; const p = await S.A.loadSheetPage(i);
  S.view = [{ idx: i, w: p.w, h: p.h }];
  for (const b of $('sheetTabs').children) b.classList.toggle('on', +b.dataset.i === i);
  if (S.sel && itemById(S.sel)?.page !== i) S.sel = null;
  showNotice(S.A.warnings); fitZoom(); buildPages(); updatePanel();
}

/* ---------- 쪽 화면 ---------- */
let io;
function fitZoom() {
  const vw = $('viewer').clientWidth - 40;
  const maxW = Math.max(...S.view.map(p => p.w));
  S.zoom = clampZoom(Math.min(vw / maxW, S.A.kind === 'xlsx' ? 1 : 1.5));
  S.fitMode = true; $('zoomFit').textContent = Math.round(S.zoom * 100) + '%';
}
const clampZoom = z => Math.max(0.15, Math.min(5, z));
function buildPages() {
  const box = $('pages'); box.innerHTML = '';
  if (io) io.disconnect();
  io = new IntersectionObserver(es => es.forEach(e => { if (e.isIntersecting) paintBg(e.target); }), { root: $('viewer'), rootMargin: '600px' });
  for (const p of S.view) {
    const el = document.createElement('div'); el.className = 'page'; el.dataset.i = p.idx;
    el.innerHTML = `<div class="bg"></div><div class="ov"></div>${S.view.length > 1 ? `<span class="pno">${p.idx + 1}</span>` : ''}`;
    box.appendChild(el); io.observe(el);
  }
  layoutPages(); renderItems();
}
function layoutPages() {
  for (const el of $('pages').children) {
    const p = S.view.find(v => v.idx === +el.dataset.i);
    el.style.width = p.w * S.zoom + 'px'; el.style.height = p.h * S.zoom + 'px';
  }
  $('zoomFit').textContent = Math.round(S.zoom * 100) + '%';
}
async function paintBg(el) {
  const i = +el.dataset.i, want = S.zoom * dpr();
  if (el._busy || (el._scale && (S.A.kind === 'hwp' || S.A.kind === 'hwpx' || (want <= el._scale * 1.15 && want >= el._scale / 1.8)))) return;
  el._busy = true;
  try {
    const node = await S.A.renderPage(i, want);
    const bg = el.querySelector('.bg'); if (!bg) return;
    const old = bg.firstChild; if (old && old._url) URL.revokeObjectURL(old._url);
    bg.innerHTML = ''; bg.appendChild(node); el._scale = want;
  } catch (e) { console.error(e); }
  finally { el._busy = false; }
}
function setZoom(z, anchor) {
  const v = $('viewer'); const old = S.zoom; z = clampZoom(z); if (Math.abs(z - old) < 1e-3) return;
  const ax = anchor ? anchor.x : v.clientWidth / 2, ay = anchor ? anchor.y : v.clientHeight / 2;
  const cx = (v.scrollLeft + ax) / old, cy = (v.scrollTop + ay) / old;
  S.zoom = z; S.fitMode = false; layoutPages(); renderItems();
  v.scrollLeft = cx * z - ax; v.scrollTop = cy * z - ay;
  clearTimeout(setZoom.t); setZoom.t = setTimeout(() => { for (const el of $('pages').children) { const r = el.getBoundingClientRect(), vr = v.getBoundingClientRect(); if (r.bottom > vr.top - 600 && r.top < vr.bottom + 600) paintBg(el); } }, 180);
}
const pageEl = i => $('pages').querySelector(`.page[data-i="${i}"]`);

/* ---------- 이미지 목록 ---------- */
async function addImages(files, at) {
  if (!files.length) return;
  for (const f of files) {
    try {
      const L = await loadImageFile(f);
      const img = { id: uid('img'), name: L.name, src: L.canvas, isJpeg: L.isJpeg, cut: { on: false, mode: 'solid', thr: 35, soft: 30, bg: null }, out: L.canvas, url: '', ver: 0, enc: null };
      img.url = await canvasUrl(img.out);
      S.images.set(img.id, img);
      renderLib();
      if (at && S.A) placeImage(img.id, at.page, at.x, at.y, true);
      else if (S.A && files.length === 1) placeImage(img.id);
    } catch (e) { await modal(e.message || String(e)); }
  }
}
async function canvasUrl(c) { const b = await new Promise(r => c.toBlob(r, 'image/png')); return URL.createObjectURL(b); }
function renderLib() {
  const lib = $('lib'); lib.innerHTML = '';
  for (const img of S.images.values()) {
    if (img.removed) continue;
    const d = document.createElement('div'); d.className = 'li checker'; d.draggable = true; d.title = img.name;
    d.innerHTML = `<img alt=""><div class="lact"><button class="btn sm" data-a="put">넣기</button><button class="btn sm ghost" data-a="rm" title="목록에서 빼기">✕</button></div>${img.cut.on ? '<span class="tag">누끼</span>' : ''}`;
    d.querySelector('img').src = img.url;
    d.addEventListener('dragstart', e => { e.dataTransfer.setData('text/x-img', img.id); e.dataTransfer.effectAllowed = 'copy'; });
    d.querySelector('[data-a=put]').onclick = () => { if (!S.A) return toast('먼저 문서를 여세요.'); placeImage(img.id); };
    d.querySelector('[data-a=rm]').onclick = async () => {
      const used = S.items.some(it => it.imgId === img.id);
      if (used && !(await confirmBox('문서에 넣은 이 이미지도 함께 지워집니다. 계속할까요?'))) return;
      if (used) { pushUndo(); S.items = S.items.filter(it => it.imgId !== img.id); S.dirty = true; renderItems(); }
      img.removed = true; renderLib(); updatePanel();
    };
    lib.appendChild(d);
  }
  $('libHint').hidden = [...S.images.values()].some(i => !i.removed);
}
function visiblePage() {
  const vr = $('viewer').getBoundingClientRect(); let best = null, bestA = -1;
  for (const el of $('pages').children) {
    const r = el.getBoundingClientRect();
    const a = Math.max(0, Math.min(r.bottom, vr.bottom) - Math.max(r.top, vr.top)) * Math.max(0, Math.min(r.right, vr.right) - Math.max(r.left, vr.left));
    if (a > bestA) { bestA = a; best = el; }
  }
  return best ? +best.dataset.i : S.view[0].idx;
}
function placeImage(imgId, page, cx, cy, centerAt) {
  const img = S.images.get(imgId); const pi = page ?? visiblePage();
  const p = S.view.find(v => v.idx === pi) || S.view[0];
  const unitPerPx = S.A.pxPerMm * 25.4 / 96;
  let w = Math.min(img.src.width * unitPerPx, p.w * 0.25), h = w * img.src.height / img.src.width;
  if (h > p.h * 0.4) { h = p.h * 0.4; w = h * img.src.width / img.src.height; }
  let x, y;
  if (cx != null) { x = cx - w / 2; y = cy - h / 2; }
  else {
    const el = pageEl(p.idx), r = el.getBoundingClientRect(), vr = $('viewer').getBoundingClientRect();
    const midY = (Math.max(r.top, vr.top) + Math.min(r.bottom, vr.bottom)) / 2;
    x = (p.w - w) / 2; y = (midY - r.top) / S.zoom - h / 2;
  }
  x = Math.max(0, Math.min(p.w - w, x)); y = Math.max(0, Math.min(p.h - h, y));
  pushUndo();
  const it = { id: uid('it'), imgId, page: p.idx, x, y, w, h };
  S.items.push(it); S.sel = it.id; S.dirty = true;
  renderItems(); updatePanel();
}

/* ---------- 배치한 이미지 ---------- */
const itemById = id => S.items.find(i => i.id === id);
function renderItems() {
  for (const pe of $('pages').children) {
    const ov = pe.querySelector('.ov'); const i = +pe.dataset.i;
    ov.querySelectorAll('.item').forEach(n => n.remove());
    for (const it of S.items) {
      if (it.page !== i) continue;
      const img = S.images.get(it.imgId);
      const d = document.createElement('div'); d.className = 'item' + (it.id === S.sel ? ' sel' : ''); d.dataset.id = it.id;
      d.innerHTML = `<img alt="" draggable="false">${it.id === S.sel ? '<i class="h" data-c="nw"></i><i class="h" data-c="ne"></i><i class="h" data-c="sw"></i><i class="h" data-c="se"></i>' : ''}`;
      d.querySelector('img').src = img.url;
      posItem(d, it); ov.appendChild(d);
    }
  }
}
function posItem(d, it) { const z = S.zoom; d.style.left = it.x * z + 'px'; d.style.top = it.y * z + 'px'; d.style.width = it.w * z + 'px'; d.style.height = it.h * z + 'px'; }

/* ---------- 스마트 가이드 ---------- */
async function snapsFor(i) {
  if (!S.snapCache.has(i)) S.snapCache.set(i, S.A.snaps(i).catch(() => ({ xs: [], ys: [] })));
  return S.snapCache.get(i);
}
function snapAxis(vals, cands, tol) {
  let best = null;
  for (const v of vals) for (const c of cands) { const d = c - v; if (Math.abs(d) <= tol && (!best || Math.abs(d) < Math.abs(best.d))) best = { d, at: c }; }
  return best;
}
function candidates(page, skipId, extra) {
  const p = S.view.find(v => v.idx === page);
  const xs = [0, p.w / 2, p.w], ys = [0, p.h / 2, p.h], strong = { xs: new Set(xs), ys: new Set(ys) };
  for (const o of S.items) if (o.page === page && o.id !== skipId) {
    for (const v of [o.x, o.x + o.w / 2, o.x + o.w]) { xs.push(v); strong.xs.add(v); }
    for (const v of [o.y, o.y + o.h / 2, o.y + o.h]) { ys.push(v); strong.ys.add(v); }
  }
  if (extra) { xs.push(...extra.xs); ys.push(...extra.ys); }
  return { xs, ys, strong };
}
function drawGuides(page, gx, gy) {
  document.querySelectorAll('.guide').forEach(n => n.remove());
  const pe = pageEl(page); if (!pe) return; const ov = pe.querySelector('.ov');
  for (const x of gx) { const g = document.createElement('div'); g.className = 'guide gx'; g.style.left = x * S.zoom + 'px'; ov.appendChild(g); }
  for (const y of gy) { const g = document.createElement('div'); g.className = 'guide gy'; g.style.top = y * S.zoom + 'px'; ov.appendChild(g); }
}

/* ---------- 끌기 / 크기 조절 ---------- */
let drag = null;
$('pages').addEventListener('pointerdown', async e => {
  const itEl = e.target.closest('.item');
  if (!itEl) { if (S.sel) { S.sel = null; renderItems(); updatePanel(); } return; }
  e.preventDefault();
  const it = itemById(itEl.dataset.id);
  if (S.sel !== it.id) { S.sel = it.id; renderItems(); updatePanel(); }
  const pr = pageEl(it.page).getBoundingClientRect();
  const handle = e.target.classList.contains('h') ? e.target.dataset.c : null;
  drag = {
    id: it.id, mode: handle ? 'resize' : 'move', corner: handle, before: JSON.stringify(S.items), moved: false,
    gx: (e.clientX - pr.left) / S.zoom - it.x, gy: (e.clientY - pr.top) / S.zoom - it.y,
    box: { ...it }, extra: null
  };
  e.target.setPointerCapture(e.pointerId);
  snapsFor(it.page).then(s => { if (drag) drag.extra = s; });
});
window.addEventListener('pointermove', e => {
  if (!drag) return;
  const it = itemById(drag.id); if (!it) return;
  let page = it.page;
  if (drag.mode === 'move' && !S.A.perSheet) {
    for (const pe of $('pages').children) { const r = pe.getBoundingClientRect(); if (e.clientY >= r.top && e.clientY <= r.bottom && e.clientX >= r.left - 60 && e.clientX <= r.right + 60) { page = +pe.dataset.i; break; } }
    if (page !== it.page) { snapsFor(page).then(s => { if (drag) drag.extra = s; }); drag.extra = null; }
  }
  const pr = pageEl(page).getBoundingClientRect();
  const px = (e.clientX - pr.left) / S.zoom, py = (e.clientY - pr.top) / S.zoom;
  const pv = S.view.find(v => v.idx === page);
  const snapping = $('snapOn').checked && !e.altKey;
  const tol = 7 / S.zoom;
  const C = snapping ? candidates(page, it.id, drag.extra) : null;
  const gx = [], gy = [];
  if (drag.mode === 'move') {
    let x = px - drag.gx, y = py - drag.gy;
    if (C) {
      const sx = snapAxis([x, x + it.w / 2, x + it.w], C.xs, tol); if (sx) { x += sx.d; gx.push(sx.at); }
      const sy = snapAxis([y, y + it.h / 2, y + it.h], C.ys, tol); if (sy) { y += sy.d; gy.push(sy.at); }
    }
    if (page !== it.page) { it.page = page; renderItems(); }
    it.x = Math.max(-it.w * 0.9, Math.min(pv.w - it.w * 0.1, x)); it.y = Math.max(-it.h * 0.9, Math.min(pv.h - it.h * 0.1, y));
  } else {
    const b = drag.box, c = drag.corner, ar = b.w / b.h;
    const ax = c.includes('w') ? b.x + b.w : b.x, ay = c.includes('n') ? b.y + b.h : b.y;
    let mx = px, my = py;
    if (C) {
      const sx = snapAxis([mx], C.xs, tol); if (sx) { mx += sx.d; gx.push(sx.at); }
      const sy = snapAxis([my], C.ys, tol); if (sy) { my += sy.d; gy.push(sy.at); }
    }
    let w = Math.max(4, Math.abs(mx - ax)), h = Math.max(4, Math.abs(my - ay));
    if ($('lockAr').checked !== e.shiftKey) {
      if (gy.length && !gx.length) w = h * ar; else h = w / ar;
      if (gx.length && gy.length) gy.length = 0;
      if (gx.length) gy.length = 0; else if (gy.length) gx.length = 0;
    }
    it.w = w; it.h = h; it.x = c.includes('w') ? ax - w : ax; it.y = c.includes('n') ? ay - h : ay;
  }
  drag.moved = true;
  const d = document.querySelector(`.item[data-id="${it.id}"]`); if (d) posItem(d, it);
  drawGuides(page, gx, gy); updateWhere();
});
window.addEventListener('pointerup', () => {
  if (!drag) return;
  if (drag.moved) { S.undo.push(drag.before); S.redo = []; S.dirty = true; updateUndoBtns(); updatePanel(); }
  drag = null; drawGuides(-1, [], []);
});

/* ---------- 실행 취소 ---------- */
function pushUndo() { S.undo.push(JSON.stringify(S.items)); if (S.undo.length > 200) S.undo.shift(); S.redo = []; updateUndoBtns(); }
function undo() { if (!S.undo.length) return; S.redo.push(JSON.stringify(S.items)); S.items = JSON.parse(S.undo.pop()); afterHistory(); }
function redo() { if (!S.redo.length) return; S.undo.push(JSON.stringify(S.items)); S.items = JSON.parse(S.redo.pop()); afterHistory(); }
function afterHistory() { if (S.sel && !itemById(S.sel)) S.sel = null; S.dirty = true; if (S.A?.perSheet && S.sel && itemById(S.sel).page !== S.sheet) S.sel = null; renderItems(); updatePanel(); updateUndoBtns(); }
function updateUndoBtns() { $('undoBtn').disabled = !S.undo.length; $('redoBtn').disabled = !S.redo.length; }

/* ---------- 속성 패널 ---------- */
const fmtMm = v => Math.round(v * 10) / 10;
function updatePanel() {
  const it = S.sel && itemById(S.sel);
  $('propSec').hidden = !it; if (!it) return;
  const img = S.images.get(it.imgId);
  $('cutOn').checked = img.cut.on; $('cutOpts').hidden = !img.cut.on;
  document.querySelector(`input[name=cutMode][value=${img.cut.mode}]`).checked = true;
  $('solidOpts').hidden = img.cut.mode !== 'solid'; $('aiHint').hidden = img.cut.mode !== 'ai';
  $('thr').value = img.cut.thr; $('soft').value = img.cut.soft;
  const bg = img.cut.bg || estimateBg(img.src); $('bgSw').style.background = `rgb(${bg})`;
  if (document.activeElement !== $('wMm')) $('wMm').value = fmtMm(it.w / S.A.pxPerMm);
  if (document.activeElement !== $('hMm')) $('hMm').value = fmtMm(it.h / S.A.pxPerMm);
  drawPreview(img); updateWhere();
}
function updateWhere() {
  const it = S.sel && itemById(S.sel); if (!it) return;
  const mm = v => fmtMm(v / S.A.pxPerMm);
  $('where').textContent = S.A.cellLabel ? `${S.A.cellLabel(it.page, it.x, it.y)} 셀 기준` : `${it.page + 1}쪽 · 왼쪽 ${mm(it.x)} · 위 ${mm(it.y)} mm`;
}
function drawPreview(img) {
  const c = $('prevCanvas'), src = img.out;
  const box = $('preview').clientWidth || 280, H = 170;
  const k = Math.min((box - 16) / src.width, H / src.height, 1.5);
  c.width = Math.max(1, Math.round(src.width * k)); c.height = Math.max(1, Math.round(src.height * k));
  const g = c.getContext('2d'); g.clearRect(0, 0, c.width, c.height); g.drawImage(src, 0, 0, c.width, c.height);
  c._k = k;
}
const curImg = () => { const it = S.sel && itemById(S.sel); return it ? S.images.get(it.imgId) : null; };
let procToken = 0;
async function reprocess(img) {
  const tok = ++procToken;
  img.ver++; img.enc = null;
  try {
    if (!img.cut.on) img.out = img.src;
    else if (img.cut.mode === 'solid') img.out = solidCutout(img.src, img.cut.bg || estimateBg(img.src), img.cut.thr, img.cut.soft);
    else {
      $('busy').hidden = false; $('busy').textContent = 'AI 모델 준비 중…';
      img.out = await aiCutout(img.src, (key, cur, total) => {
        if (/model|wasm/i.test(key) && total) $('busy').textContent = `AI 모델 내려받는 중 ${Math.round(cur / total * 100)}%`;
        else $('busy').textContent = '배경 지우는 중…';
      });
    }
  } catch (e) {
    console.error(e); img.cut.on = false; img.out = img.src;
    await modal('누끼 처리에 실패했습니다: ' + (e.message || e));
  } finally { if (tok === procToken) $('busy').hidden = true; }
  const old = img.url; img.url = await canvasUrl(img.out); URL.revokeObjectURL(old);
  S.dirty = true; renderItems(); renderLib(); if (curImg() === img) { drawPreview(img); updatePanel(); }
}
let reT;
const reprocessSoon = img => { clearTimeout(reT); reT = setTimeout(() => reprocess(img), 120); };
$('cutOn').onchange = e => { const img = curImg(); if (!img) return; img.cut.on = e.target.checked; updatePanel(); reprocess(img); };
document.querySelectorAll('input[name=cutMode]').forEach(r => r.onchange = () => { const img = curImg(); if (!img) return; img.cut.mode = r.value; updatePanel(); reprocess(img); });
$('thr').oninput = e => { const img = curImg(); if (!img) return; img.cut.thr = +e.target.value; reprocessSoon(img); };
$('soft').oninput = e => { const img = curImg(); if (!img) return; img.cut.soft = +e.target.value; reprocessSoon(img); };
$('autoBg').onclick = () => { const img = curImg(); if (!img) return; img.cut.bg = null; updatePanel(); reprocess(img); };
let picking = false;
$('pickBtn').onclick = () => { picking = !picking; $('pickBtn').classList.toggle('on', picking); $('prevCanvas').style.cursor = picking ? 'crosshair' : ''; if (picking) toast('미리보기에서 지울 배경 부분을 누르세요.'); };
$('prevCanvas').addEventListener('click', e => {
  const img = curImg(); if (!picking || !img) return;
  const r = e.target.getBoundingClientRect(), k = e.target._k || 1;
  const sc = e.target.width / r.width; img.cut.bg = pickColor(img.src, Math.round((e.clientX - r.left) * sc / k), Math.round((e.clientY - r.top) * sc / k));
  picking = false; $('pickBtn').classList.remove('on'); e.target.style.cursor = '';
  updatePanel(); reprocess(img);
});
function setSizeMm(which, val) {
  const it = itemById(S.sel); if (!it || !(val > 0)) return;
  pushUndo(); const ar = it.w / it.h; const v = val * S.A.pxPerMm;
  const cx = it.x, cy = it.y;
  if (which === 'w') { it.w = v; if ($('lockAr').checked) it.h = v / ar; } else { it.h = v; if ($('lockAr').checked) it.w = v * ar; }
  it.x = cx; it.y = cy; S.dirty = true; renderItems(); updatePanel();
}
$('wMm').onchange = e => setSizeMm('w', +e.target.value);
$('hMm').onchange = e => setSizeMm('h', +e.target.value);
$('dupBtn').onclick = () => duplicate();
$('delBtn').onclick = () => removeSel();
$('centerBtn').onclick = () => { const it = itemById(S.sel); if (!it) return; pushUndo(); const p = S.view.find(v => v.idx === it.page); it.x = (p.w - it.w) / 2; S.dirty = true; renderItems(); updatePanel(); };
function duplicate() { const it = itemById(S.sel); if (!it) return; pushUndo(); const n = { ...it, id: uid('it'), x: it.x + 12, y: it.y + 12 }; S.items.push(n); S.sel = n.id; S.dirty = true; renderItems(); updatePanel(); }
function removeSel() { if (!S.sel) return; pushUndo(); S.items = S.items.filter(i => i.id !== S.sel); S.sel = null; S.dirty = true; renderItems(); updatePanel(); }

/* ---------- 저장 ---------- */
async function getImage(imgId) {
  const img = S.images.get(imgId);
  if (!img.enc) img.enc = { key: img.id + ':' + img.ver, name: img.name, ...(await encodeForInsert(img.out, !img.cut.on && img.isJpeg)) };
  return img.enc;
}
async function save() {
  if (!S.A) return;
  if (!S.items.length && !(await confirmBox('넣은 이미지가 없습니다. 그래도 저장할까요?'))) return;
  const btn = $('saveBtn'); btn.disabled = true;
  try {
    toast('저장 중…', 0);
    const items = S.items.map(i => ({ ...i }));
    const res = await S.A.save(items, getImage, m => toast(m, 0));
    const blob = res instanceof Blob ? res : res.blob;
    const warnings = res instanceof Blob ? [] : (res.warnings || []);
    const base = S.fileName.replace(/\.[^.]+$/, '');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = `${base}_이미지삽입.${S.A.ext}`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 60000);
    S.dirty = false;
    toast('저장했습니다: ' + a.download, 4000);
    if (warnings.length) await modal(warnings.join('\n'));
  } catch (e) { console.error(e); $('toast').classList.remove('on'); await modal('저장하지 못했습니다. ' + (e.message || e)); }
  finally { btn.disabled = false; }
}

/* ---------- 입력 연결 ---------- */
$('docInput').onchange = e => { openDoc(e.target.files[0]); e.target.value = ''; };
$('docInput2').onchange = e => { openDoc(e.target.files[0]); e.target.value = ''; };
$('imgInput').onchange = e => { addImages([...e.target.files]); e.target.value = ''; };
$('saveBtn').onclick = save;
$('undoBtn').onclick = undo; $('redoBtn').onclick = redo;
$('zoomIn').onclick = () => setZoom(S.zoom * 1.25);
$('zoomOut').onclick = () => setZoom(S.zoom / 1.25);
$('zoomFit').onclick = () => { if (!S.A) return; fitZoom(); layoutPages(); renderItems(); for (const el of $('pages').children) paintBg(el); };
$('viewer').addEventListener('wheel', e => {
  if (!S.A || !(e.ctrlKey || e.metaKey)) return; e.preventDefault();
  const r = $('viewer').getBoundingClientRect(); setZoom(S.zoom * Math.exp(-e.deltaY * 0.002), { x: e.clientX - r.left, y: e.clientY - r.top });
}, { passive: false });
window.addEventListener('resize', () => { if (S.A && S.fitMode) { fitZoom(); layoutPages(); renderItems(); } });

const isDocName = n => /\.(pdf|xlsx|xlsm|xls|hwp|hwpx)$/i.test(n);
$('viewer').addEventListener('dragover', e => { e.preventDefault(); $('viewer').classList.add('dropping'); });
$('viewer').addEventListener('dragleave', e => { if (e.target === $('viewer')) $('viewer').classList.remove('dropping'); });
$('viewer').addEventListener('drop', async e => {
  e.preventDefault(); $('viewer').classList.remove('dropping');
  const at = dropPoint(e);
  const libId = e.dataTransfer.getData('text/x-img');
  if (libId) { if (S.A && at) placeImage(libId, at.page, at.x, at.y, true); return; }
  const files = [...e.dataTransfer.files];
  const doc = files.find(f => isDocName(f.name));
  if (doc) return openDoc(doc);
  const imgs = files.filter(f => /^image\//.test(f.type) || /\.(png|jpe?g|gif|webp|bmp|svg|heic|heif)$/i.test(f.name));
  if (imgs.length) addImages(imgs, S.A ? at : null);
});
function dropPoint(e) {
  for (const pe of $('pages').children) {
    const r = pe.getBoundingClientRect();
    if (e.clientX >= r.left && e.clientX <= r.right && e.clientY >= r.top && e.clientY <= r.bottom)
      return { page: +pe.dataset.i, x: (e.clientX - r.left) / S.zoom, y: (e.clientY - r.top) / S.zoom };
  }
  return null;
}
document.querySelector('.panel').addEventListener('dragover', e => e.preventDefault());
document.querySelector('.panel').addEventListener('drop', e => { e.preventDefault(); const f = [...e.dataTransfer.files].filter(f => !isDocName(f.name)); addImages(f); });

document.addEventListener('keydown', e => {
  const tag = (e.target.tagName || '').toLowerCase();
  if (tag === 'input' || tag === 'textarea' || !$('modal').hidden) return;
  const mod = e.ctrlKey || e.metaKey;
  if (mod && e.key.toLowerCase() === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
  if (mod && e.key.toLowerCase() === 'y') { e.preventDefault(); redo(); return; }
  if (mod && e.key.toLowerCase() === 's') { e.preventDefault(); save(); return; }
  const it = S.sel && itemById(S.sel); if (!it) return;
  if (mod && e.key.toLowerCase() === 'd') { e.preventDefault(); duplicate(); return; }
  if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); removeSel(); return; }
  if (e.key === 'Escape') { S.sel = null; renderItems(); updatePanel(); return; }
  const step = e.shiftKey ? 10 : 1, mv = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[e.key];
  if (mv) {
    e.preventDefault();
    if (!e.repeat) pushUndo();
    it.x += mv[0]; it.y += mv[1]; S.dirty = true;
    const d = document.querySelector(`.item[data-id="${it.id}"]`); if (d) posItem(d, it); updateWhere();
  }
});
window.addEventListener('beforeunload', e => { if (S.dirty && S.items.length) { e.preventDefault(); e.returnValue = ''; } });

// 테스트·자동화용 최소 노출
window.__app = { S, openDoc, addImages, placeImage, save, reprocess, setZoom };
