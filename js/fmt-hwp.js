// HWP / HWPX: rhwp(WASM)로 렌더링하고, 쪽 기준 절대 위치의 "글 앞으로" 그림을 넣어 같은 형식으로 저장
const HU = 75; // 1px(96dpi) = 75 HWPUNIT
let rhwp = null;

async function ensureRhwp(onStatus) {
  if (rhwp) return rhwp;
  onStatus && onStatus('한글 문서 엔진을 불러오는 중… (처음 한 번, 약 28MB)');
  const mod = await import('../vendor/rhwp/rhwp.js');
  let ctx = null, last = '';
  globalThis.measureTextWidth = (font, text) => {
    if (!ctx) ctx = document.createElement('canvas').getContext('2d');
    if (font !== last) { ctx.font = font; last = font; }
    return ctx.measureText(text).width;
  };
  await mod.default({ module_or_path: new URL('../vendor/rhwp/rhwp_bg.wasm', import.meta.url).href });
  rhwp = mod;
  return rhwp;
}

export async function openHwp(file, onStatus) {
  const isX = /\.hwpx$/i.test(file.name);
  const bytes = new Uint8Array(await file.arrayBuffer());
  const m = await ensureRhwp(onStatus);
  let doc;
  try { doc = new m.HwpDocument(bytes); }
  catch (e) {
    throw new Error('한글 문서를 열 수 없습니다. 암호가 걸렸거나 배포용 문서일 수 있습니다. 한글에서 일반 문서로 다시 저장한 뒤 올려 주세요.' + (e && e.message ? ` (${e.message})` : ''));
  }
  const n = doc.pageCount();
  if (!n) throw new Error('문서에 쪽이 없습니다.');
  const infos = [];
  for (let i = 0; i < n; i++) infos.push(JSON.parse(doc.getPageInfo(i)));
  const svgCache = new Map();
  const getSvg = i => { if (!svgCache.has(i)) svgCache.set(i, doc.renderPageSvg(i)); return svgCache.get(i); };

  return {
    kind: isX ? 'hwpx' : 'hwp', ext: isX ? 'hwpx' : 'hwp',
    mime: isX ? 'application/hwp+zip' : 'application/x-hwp',
    warnings: [], pxPerMm: 96 / 25.4,
    pages: infos.map(p => ({ w: p.width, h: p.height })),
    async renderPage(i) {
      // SVG는 <img>로 표시(스크립트 실행 없음)
      const url = URL.createObjectURL(new Blob([getSvg(i)], { type: 'image/svg+xml' }));
      const im = new Image();
      im.decoding = 'async';
      await new Promise((res, rej) => { im.onload = res; im.onerror = () => rej(new Error('쪽 그림 생성 실패')); im.src = url; });
      im._url = url;
      return im;
    },
    async snaps(i) {
      const p = infos[i];
      const xs = [p.bodyLeft, p.bodyRight], ys = [];
      if (p.headerArea) ys.push(p.headerArea.y + p.headerArea.height);
      if (p.footerArea) ys.push(p.footerArea.y);
      for (const c of p.columns || []) xs.push(c.x, c.x + c.width);
      // 글자 줄: 같은 기준선끼리 묶어 줄 시작·끝·기준선
      const lines = new Map();
      for (const mm of getSvg(i).matchAll(/<text x="([\d.]+)" y="([\d.]+)"[^>]*?font-size="([\d.]+)"/g)) {
        const x = +mm[1], y = Math.round(+mm[2] * 2) / 2, fs = +mm[3];
        const L = lines.get(y) || { a: x, b: x, fs };
        L.a = Math.min(L.a, x); L.b = Math.max(L.b, x + fs * 0.9); lines.set(y, L);
      }
      for (const [y, L] of lines) { ys.push(y, y - L.fs * 0.85); xs.push(L.a, L.b); }
      const d = a => [...new Set(a.filter(Number.isFinite).map(v => Math.round(v * 2) / 2))].sort((p, q) => p - q);
      return { xs: d(xs), ys: d(ys) };
    },
    async save(items, getImage, onStatus) {
      const out = new m.HwpDocument(bytes);
      const pagesBefore = out.pageCount();
      const failed = [];
      let k = 0;
      for (const it of items) {
        k++; onStatus && onStatus(`그림 넣는 중… (${k}/${items.length})`);
        const img = await getImage(it.imgId);
        if (!placeOne(out, it, img)) failed.push(it);
      }
      if (failed.length) throw new Error(`${failed.length}개 그림의 위치를 문서에 고정하지 못했습니다(해당 쪽에 기준이 될 문단이 없음). 그 그림을 본문 글자가 있는 쪽으로 옮겨 다시 시도해 주세요.`);
      const rep = isX ? out.exportHwpxWithReport() : out.exportHwpWithReport();
      const loss = JSON.parse(rep.contentLoss() || '{}');
      const data = rep.takeBytes();
      const warnings = [];
      if (loss.count > 0) warnings.push(`저장 엔진이 원본 일부 요소(${loss.count}개)를 그대로 옮기지 못했다고 보고했습니다. 한글에서 열어 확인해 주세요.`);
      try {
        const re = new m.HwpDocument(data);
        if (re.pageCount() !== pagesBefore) warnings.push(`저장 후 쪽 수가 달라졌습니다(${pagesBefore} → ${re.pageCount()}). 한글에서 열어 확인해 주세요.`);
        re.free && re.free();
      } catch (e) { warnings.push('저장된 파일을 다시 읽는 검사에 실패했습니다. 한글에서 열어 확인해 주세요.'); }
      out.free && out.free();
      return { blob: new Blob([data], { type: this.mime }), warnings };
    }
  };
}

// 그림 하나 넣기: 대상 쪽 위의 문단을 기준으로 삽입 → 쪽 기준 절대 위치·글 앞으로 → 렌더로 쪽·위치 검증
function placeOne(doc, it, img) {
  const info = JSON.parse(doc.getPageInfo(it.page));
  const top = info.headerArea ? info.headerArea.y + info.headerArea.height : 60;
  const bottom = info.footerArea ? info.footerArea.y : info.height - 60;
  const probes = [[it.x + 2, it.y + 2], [it.x + it.w / 2, it.y + it.h / 2]];
  for (let y = top + 4; y < bottom; y += 24) probes.push([info.bodyLeft + 4, y]);
  const tried = new Set();
  for (const [px, py] of probes) {
    let h;
    try { h = JSON.parse(doc.hitTest(it.page, px, py)); } catch (e) { continue; }
    if (!h || h.sectionIndex == null) continue;
    const inBox = Array.isArray(h.cellPath) && h.cellPath.length;
    const para = inBox ? (h.parentParaIndex ?? h.paragraphIndex) : h.paragraphIndex;
    const off = inBox ? 0 : (h.charOffset || 0);
    const key = `${h.sectionIndex}:${para}:${off}`;
    if (tried.has(key)) continue; tried.add(key);
    let r;
    try {
      r = JSON.parse(doc.insertPicture(h.sectionIndex, para, off, '', img.bytes,
        Math.round(it.w * HU), Math.round(it.h * HU), img.w, img.h, img.ext, img.name || '그림'));
    } catch (e) { continue; }
    if (!r || !r.ok) continue;
    doc.setPictureProperties(h.sectionIndex, r.paraIdx, r.controlIdx, JSON.stringify({
      treatAsChar: false, textWrap: 'InFrontOfText', horzRelTo: 'Paper', vertRelTo: 'Paper',
      horzAlign: 'Left', vertAlign: 'Top', allowOverlap: true, restrictInPage: false,
      horzOffset: Math.round(it.x * HU), vertOffset: Math.round(it.y * HU),
      width: Math.round(it.w * HU), height: Math.round(it.h * HU)
    }));
    if (onPage(doc, it)) return true;
    try { doc.deletePictureControl(h.sectionIndex, r.paraIdx, r.controlIdx); } catch (e) { }
  }
  return false;
}
function onPage(doc, it) {
  const svg = doc.renderPageSvg(it.page);
  for (const mm of svg.matchAll(/<image x="([-\d.]+)" y="([-\d.]+)" width="([\d.]+)" height="([\d.]+)"/g)) {
    if (Math.abs(+mm[1] - it.x) < 1.5 && Math.abs(+mm[2] - it.y) < 1.5 && Math.abs(+mm[3] - it.w) < 1.5 && Math.abs(+mm[4] - it.h) < 1.5) return true;
  }
  return false;
}
