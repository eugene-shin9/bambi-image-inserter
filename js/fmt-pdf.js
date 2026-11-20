// PDF: pdf.js로 미리보기, pdf-lib로 원본 위에 이미지 그리기
import * as pdfjs from '../vendor/pdfjs/pdf.min.mjs';
pdfjs.GlobalWorkerOptions.workerSrc = new URL('../vendor/pdfjs/pdf.worker.min.mjs', import.meta.url).href;
const PDFLib = () => window.PDFLib;

export async function openPdf(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const warnings = [];
  // 1) 저장 가능 여부 먼저 확인 (암호·보안 PDF는 pdf-lib이 저장하지 못함)
  let libDoc;
  try {
    libDoc = await PDFLib().PDFDocument.load(bytes, { updateMetadata: false });
  } catch (e) {
    if (/encrypt/i.test(String(e && (e.name + e.message)))) {
      throw new Error('보안(암호)이 걸린 PDF는 저장할 수 없습니다. PDF 뷰어에서 "인쇄 → PDF로 저장"으로 보안 없는 사본을 만든 뒤 올려 주세요.');
    }
    throw new Error('PDF를 읽을 수 없습니다: ' + (e.message || e));
  }
  try {
    const form = libDoc.getForm();
    if (form.getFields().some(f => f.constructor && /Signature/.test(f.constructor.name))) {
      warnings.push('전자서명이 들어 있는 PDF입니다. 이미지를 넣어 저장하면 기존 전자서명은 무효가 됩니다.');
    }
  } catch (e) { /* 양식 없음 */ }

  const task = pdfjs.getDocument({
    data: bytes.slice(),
    cMapUrl: new URL('../vendor/pdfjs/cmaps/', import.meta.url).href,
    cMapPacked: true,
    standardFontDataUrl: new URL('../vendor/pdfjs/standard_fonts/', import.meta.url).href,
    isEvalSupported: false
  });
  let pdf;
  try { pdf = await task.promise; }
  catch (e) {
    if (e && e.name === 'PasswordException') throw new Error('열기 암호가 걸린 PDF입니다. 암호를 푼 사본을 올려 주세요.');
    throw new Error('PDF를 열 수 없습니다: ' + (e.message || e));
  }
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const p = await pdf.getPage(i);
    const vp = p.getViewport({ scale: 1 });
    pages.push({ w: vp.width, h: vp.height, page: p, vp });
  }
  const snapCache = new Map();

  return {
    kind: 'pdf', ext: 'pdf', mime: 'application/pdf', warnings,
    pxPerMm: 72 / 25.4,
    pages: pages.map(p => ({ w: p.w, h: p.h })),
    async renderPage(i, scale) {
      const { page } = pages[i];
      const vp = page.getViewport({ scale });
      const c = document.createElement('canvas');
      c.width = Math.ceil(vp.width); c.height = Math.ceil(vp.height);
      const ctx = c.getContext('2d');
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
      await page.render({ canvasContext: ctx, viewport: vp }).promise;
      return c;
    },
    async snaps(i) {
      if (snapCache.has(i)) return snapCache.get(i);
      const { page, vp } = pages[i];
      const xs = [], ys = [];
      try {
        const tc = await page.getTextContent();
        for (const it of tc.items) {
          if (!it.str || !it.str.trim()) continue;
          const m = pdfjs.Util.transform(vp.transform, it.transform);
          const fh = Math.hypot(m[2], m[3]);
          const x = m[4], y = m[5];
          if (Math.abs(m[1]) > 0.01 || Math.abs(m[2]) > 0.01) continue; // 기울어진 글자는 제외
          xs.push(x, x + it.width * vp.scale);
          ys.push(y, y - fh);
        }
      } catch (e) { /* 글자 정보 없음 */ }
      const res = { xs: dedupe(xs), ys: dedupe(ys) };
      snapCache.set(i, res);
      return res;
    },
    async save(items, getImage) {
      const doc = await PDFLib().PDFDocument.load(bytes, { updateMetadata: false });
      const embedded = new Map();
      for (const it of items) {
        const img = await getImage(it.imgId);
        let emb = embedded.get(img.key);
        if (!emb) {
          emb = img.ext === 'jpg' ? await doc.embedJpg(img.bytes) : await doc.embedPng(img.bytes);
          embedded.set(img.key, emb);
        }
        const { vp } = pages[it.page];
        const page = doc.getPage(it.page);
        // 화면상 좌하단·우하단·좌상단 → PDF 사용자 좌표 (쪽 회전·CropBox 원점 자동 반영)
        const P0 = vp.convertToPdfPoint(it.x, it.y + it.h);
        const P1 = vp.convertToPdfPoint(it.x + it.w, it.y + it.h);
        const P2 = vp.convertToPdfPoint(it.x, it.y);
        const width = Math.hypot(P1[0] - P0[0], P1[1] - P0[1]);
        const height = Math.hypot(P2[0] - P0[0], P2[1] - P0[1]);
        const ang = Math.atan2(P1[1] - P0[1], P1[0] - P0[0]) * 180 / Math.PI;
        page.drawImage(emb, { x: P0[0], y: P0[1], width, height, rotate: PDFLib().degrees(ang) });
      }
      const out = await doc.save();
      return new Blob([out], { type: 'application/pdf' });
    }
  };
}

function dedupe(arr) {
  const s = new Set(arr.map(v => Math.round(v * 2) / 2));
  return [...s].sort((a, b) => a - b).slice(0, 4000);
}
