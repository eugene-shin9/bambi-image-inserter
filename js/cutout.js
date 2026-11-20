// 이미지 불러오기 · 누끼(단색 배경 / AI) · 삽입용 인코딩
const MAX_DIM = 3000;

export async function loadImageFile(file) {
  const name = file.name || '이미지';
  const lower = name.toLowerCase();
  if (/\.(heic|heif)$/.test(lower) || /hei[cf]/.test(file.type)) {
    throw new Error('HEIC 사진은 지원하지 않습니다. 사진 앱에서 JPG로 내보낸 뒤 올려 주세요.');
  }
  let bmp;
  try {
    bmp = await createImageBitmap(file, { imageOrientation: 'from-image' });
  } catch (e) {
    bmp = await new Promise((res, rej) => {
      const url = URL.createObjectURL(file);
      const im = new Image();
      im.onload = () => { URL.revokeObjectURL(url); res(im); };
      im.onerror = () => { URL.revokeObjectURL(url); rej(new Error(`"${name}"을(를) 이미지로 읽을 수 없습니다.`)); };
      im.src = url;
    });
  }
  const w0 = bmp.width || bmp.naturalWidth, h0 = bmp.height || bmp.naturalHeight;
  if (!w0 || !h0) throw new Error(`"${name}"의 크기를 알 수 없습니다.`);
  const k = Math.min(1, MAX_DIM / Math.max(w0, h0));
  const c = document.createElement('canvas');
  c.width = Math.max(1, Math.round(w0 * k));
  c.height = Math.max(1, Math.round(h0 * k));
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
  if (bmp.close) bmp.close();
  const isJpeg = /jpe?g/.test(file.type) || /\.jpe?g$/.test(lower);
  return { canvas: c, isJpeg, name, file };
}

// 테두리 픽셀 색의 중앙값으로 배경색 추정
export function estimateBg(canvas) {
  const { width: w, height: h } = canvas;
  const d = canvas.getContext('2d').getImageData(0, 0, w, h).data;
  const rs = [], gs = [], bs = [];
  const step = Math.max(1, Math.floor((w + h) / 400));
  const push = (x, y) => { const i = (y * w + x) * 4; if (d[i + 3] < 16) return; rs.push(d[i]); gs.push(d[i + 1]); bs.push(d[i + 2]); };
  for (let x = 0; x < w; x += step) { push(x, 0); push(x, h - 1); }
  for (let y = 0; y < h; y += step) { push(0, y); push(w - 1, y); }
  if (!rs.length) return [255, 255, 255];
  const med = a => { a.sort((p, q) => p - q); return a[a.length >> 1]; };
  return [med(rs), med(gs), med(bs)];
}

export function pickColor(canvas, x, y) {
  const d = canvas.getContext('2d').getImageData(Math.max(0, x - 1), Math.max(0, y - 1), 3, 3).data;
  let r = 0, g = 0, b = 0, n = 0;
  for (let i = 0; i < d.length; i += 4) { r += d[i]; g += d[i + 1]; b += d[i + 2]; n++; }
  return [Math.round(r / n), Math.round(g / n), Math.round(b / n)];
}

// 단색 배경 제거: 배경색과의 색 거리로 알파 결정 + 경계 색 보정(배경색 빼기)
// strength 0~100: 클수록 배경에 가까운 색을 더 많이 지움 / soft 0~100: 경계 부드러움
export function solidCutout(src, bg, strength, soft) {
  const { width: w, height: h } = src;
  const out = document.createElement('canvas');
  out.width = w; out.height = h;
  const sctx = src.getContext('2d');
  const img = sctx.getImageData(0, 0, w, h);
  const d = img.data;
  const t0 = 4 + strength * 2.2;           // 이 거리 이하 = 완전 투명
  const t1 = t0 + 6 + soft * 1.2;          // 이 거리 이상 = 완전 불투명
  const [br, bgc, bb] = bg;
  for (let i = 0; i < d.length; i += 4) {
    const r = d[i], g = d[i + 1], b = d[i + 2];
    const dist = Math.sqrt((r - br) ** 2 + (g - bgc) ** 2 + (b - bb) ** 2);
    let a = dist <= t0 ? 0 : dist >= t1 ? 1 : (dist - t0) / (t1 - t0);
    if (a > 0 && a < 1) {
      // 반투명 경계에서 배경색 성분 제거 → 흰 테두리 줄이기
      d[i] = clamp((r - (1 - a) * br) / a);
      d[i + 1] = clamp((g - (1 - a) * bgc) / a);
      d[i + 2] = clamp((b - (1 - a) * bb) / a);
    }
    d[i + 3] = Math.round(d[i + 3] * a);
  }
  out.getContext('2d').putImageData(img, 0, 0);
  return out;
}
const clamp = v => v < 0 ? 0 : v > 255 ? 255 : Math.round(v);

// AI 누끼 (브라우저 안에서 실행, 모델은 이 사이트에 함께 들어 있음)
let bgMod = null;
export async function aiCutout(src, onProgress) {
  if (!bgMod) bgMod = await import('../vendor/bgremoval/bgremoval.mjs');
  const blob = await new Promise(r => src.toBlob(r, 'image/png'));
  const publicPath = new URL('../vendor/bgremoval/data/', import.meta.url).href;
  const run = proxy => bgMod.removeBackground(blob, {
    publicPath, model: 'small', proxyToWorker: proxy, output: { format: 'image/png' },
    progress: (key, cur, total) => onProgress && onProgress(key, cur, total)
  });
  let res;
  try { res = await run(true); } catch (e) { console.warn('worker 실행 실패, 기본 실행으로 재시도', e); res = await run(false); }
  const bmp = await createImageBitmap(res);
  const c = document.createElement('canvas');
  c.width = src.width; c.height = src.height;
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
  return c;
}

// 문서에 넣을 바이트: 누끼면 PNG, 원본 JPEG이면 JPEG, 그 외 PNG
export async function encodeForInsert(canvas, preferJpeg) {
  const type = preferJpeg ? 'image/jpeg' : 'image/png';
  const blob = await new Promise(r => canvas.toBlob(r, type, 0.92));
  return { bytes: new Uint8Array(await blob.arrayBuffer()), ext: preferJpeg ? 'jpg' : 'png', mime: type, w: canvas.width, h: canvas.height };
}
