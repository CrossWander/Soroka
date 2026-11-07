// App and tray icon, drawn in code: a magpie (soroka) on the app's purple tile, optionally a red badge like Teams'.
// Pure Node (no Electron), so the build can render build/icon.png with it too.
import zlib from 'node:zlib';

type RGBA = [number, number, number, number];
type Shape = (x: number, y: number) => boolean;
const PURPLE: RGBA = [91, 95, 199, 255], GREY: RGBA = [138, 136, 134, 255], WHITE: RGBA = [255, 255, 255, 255], RED: RGBA = [209, 52, 56, 255];
const INK: RGBA = [22, 22, 30, 255]; // the magpie's black

const rrect = (x0: number, y0: number, x1: number, y1: number, r: number): Shape => (x, y) => {
  const dx = Math.max(x0 + r - x, 0, x - (x1 - r)), dy = Math.max(y0 + r - y, 0, y - (y1 - r));
  return x >= x0 && x <= x1 && y >= y0 && y <= y1 && dx * dx + dy * dy <= r * r;
};
const circle = (cx: number, cy: number, r: number): Shape => (x, y) => (x - cx) ** 2 + (y - cy) ** 2 <= r * r;
const ellipse = (cx: number, cy: number, rx: number, ry: number, deg = 0): Shape => {
  const c = Math.cos((deg * Math.PI) / 180), s = Math.sin((deg * Math.PI) / 180);
  return (x, y) => { const dx = x - cx, dy = y - cy; return ((dx * c + dy * s) / rx) ** 2 + ((dy * c - dx * s) / ry) ** 2 <= 1; };
};
const both = (a: Shape, b: Shape): Shape => (x, y) => a(x, y) && b(x, y);
const any = (...shapes: Shape[]): Shape => (x, y) => shapes.some(s => s(x, y));
const tri = (ax: number, ay: number, bx: number, by: number, cx: number, cy: number): Shape => (x, y) => {
  const side = (px: number, py: number, qx: number, qy: number) => (x - qx) * (py - qy) - (px - qx) * (y - qy);
  const d = [side(ax, ay, bx, by), side(bx, by, cx, cy), side(cx, cy, ax, ay)];
  return !(d.some(v => v < 0) && d.some(v => v > 0));
};
// Vyshyvanka embroidery along a sleeve whose axis goes through (cx, cy) at `deg`: a chain of red rhombuses with black
// hearts, in cross-stitch. Small icons get fewer, coarser stitches; the tray only a dotted red line. Returns [red, black].
const embroidery = (cx: number, cy: number, deg: number, size: number, k: number): [Shape, Shape] => {
  const st = Math.max(1, Math.floor(size / 64)) / k, r = size < 32 ? 0 : size < 48 ? 1 : 2, period = 2 * r + 2;
  const c = Math.cos((deg * Math.PI) / 180), s = Math.sin((deg * Math.PI) / 180);
  const cell = (x: number, y: number) => {
    const dx = x - cx, dy = y - cy, gu = Math.round((dx * c + dy * s) / st), dv = Math.round((dy * c - dx * s) / st);
    const du = (((gu + r + 1) % period) + period) % period - (r + 1);
    return { du, dv, d: Math.abs(du) + Math.abs(dv) };
  };
  return [
    (x, y) => cell(x, y).d === r,
    (x, y) => { const { du, dv, d } = cell(x, y); return r > 0 && (d === 0 || (dv === 0 && du === -(r + 1))); },
  ];
};
// 3x5 pixel font for the badge
const FONT: Record<string, string> = {
  0: '111101101101111', 1: '010110010010111', 2: '111001111100111', 3: '111001111001111', 4: '101101111001001',
  5: '111100111001111', 6: '111100111101111', 7: '111001010010010', 8: '111101111101111', 9: '111101111001111',
  '+': '000010111010000', '!': '010010010000010',
};
const label = (s: string, cx: number, cy: number, cell: number): Shape => {
  const w = s.length * 4 - 1, x0 = cx - (w * cell) / 2, y0 = cy - 2.5 * cell;
  return (x, y) => {
    const gx = Math.floor((x - x0) / cell), gy = Math.floor((y - y0) / cell);
    return gy >= 0 && gy < 5 && gx >= 0 && gx < w && gx % 4 < 3 && FONT[s[Math.floor(gx / 4)]]?.[gy * 3 + (gx % 4)] === '1';
  };
};

/** RGBA pixels of a size×size icon; `badge` is "1".."9", "9+" or "!"; `grey` while not connected. */
export function draw(size: number, { badge, grey }: { badge?: string; grey?: boolean } = {}) {
  const tile = grey ? GREY : PURPLE, k = size / 32;
  const m = size > 32 ? 1 : 0; // no margin at tray sizes: there it only made the tile smaller than its neighbours and blurred its edge
  // the magpie faces left, so a badge in the bottom right corner covers only its tail
  const body = ellipse(16, 17.5, 7.2, 4.9, 22), wing = ellipse(18, 15.8, 6.2, 2.4, 22);
  // the wing is a white vyshyvanka sleeve inside a thin black rim (at least ¾ px), which parts it from the tile and the belly
  const rim = Math.max(0.6, 0.75 / k), sleeve = both(ellipse(16, 17.5, 7.2 - rim, 4.9 - rim, 22), ellipse(18, 15.8, 6.2 - rim, 2.4 - rim, 22));
  const [red, black] = embroidery(18, 15.8, 22, size, k);
  const layers: [Shape, RGBA | null][] = [ // design grid 32×32, later layers on top; null cuts a hole
    [rrect(m, m, 32 - m, 32 - m, 7), tile],
    [any(tri(19.5, 18, 21, 21.5, 29.5, 29), tri(19.5, 18, 29.5, 29, 30.5, 26.5)), INK], // the long tail, down to the right
    [any(body, circle(9.2, 10.6, 3.9), ellipse(11.6, 13.4, 3.4, 3, -35)), INK], // body, head and neck
    [tri(6, 9.4, 1.8, 11, 6, 12.5), INK], // beak
    [sleeve, WHITE],
    [both(sleeve, red), RED],
    [both(sleeve, black), INK],
    [both(both(body, ellipse(13.4, 20.6, 5.2, 3.2, 12)), (x, y) => !wing(x, y)), WHITE], // the white belly, below the rim
    [circle(8.2, 9.8, 0.75), WHITE], // eye
  ];
  if (badge) { // like Teams: a pill in the bottom right corner, cut out of the tile; the digits on whole pixels stay sharp
    const c = Math.max(1, Math.floor(size / 16)), h = 9 * c, w = Math.max(h, (badge.length * 4 - 1) * c + 4 * c);
    const x0 = (size - w) / k, y0 = (size - h) / k, g = c / k;
    layers.push([rrect(x0 - g, y0 - g, 32 + g, 32 + g, h / 2 / k + g), null], [rrect(x0, y0, 32, 32, h / 2 / k), RED],
      [label(badge, x0 + w / 2 / k, y0 + h / 2 / k, g), WHITE]);
  }

  const N = 4, px = Buffer.alloc(size * size * 4);
  for (let y = 0; y < size; y++) for (let x = 0; x < size; x++) {
    let r = 0, g = 0, b = 0, hits = 0;
    for (let sy = 0; sy < N; sy++) for (let sx = 0; sx < N; sx++) { // 4×4 supersampling for smooth edges
      const u = (x + (sx + 0.5) / N) / k, v = (y + (sy + 0.5) / N) / k;
      for (let i = layers.length - 1; i >= 0; i--) if (layers[i][0](u, v)) { const c = layers[i][1]; if (c) { r += c[0]; g += c[1]; b += c[2]; hits++; } break; }
    }
    const i = (y * size + x) * 4;
    if (hits) { px[i] = r / hits; px[i + 1] = g / hits; px[i + 2] = b / hits; px[i + 3] = (255 * hits) / (N * N); }
  }
  return px;
}

export function png(size: number, opts?: Parameters<typeof draw>[1]) {
  const px = draw(size, opts), raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) px.copy(raw, y * (size * 4 + 1) + 1, y * size * 4, (y + 1) * size * 4); // filter byte 0 per row
  const chunk = (type: string, data: Buffer) => {
    const head = Buffer.alloc(4); head.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type), data]);
    const crc = Buffer.alloc(4); crc.writeUInt32BE(zlib.crc32(body));
    return Buffer.concat([head, body, crc]);
  };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(size, 0); ihdr.writeUInt32BE(size, 4); ihdr[8] = 8; ihdr[9] = 6; // 8-bit RGBA
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

// A 32-bit bitmap entry of an .ico: rows bottom-up in BGRA, then an AND mask left empty (the alpha channel decides).
// Not every reader of .ico files takes PNG entries below 256 px, so those sizes go in as bitmaps.
function dib(size: number) {
  const px = draw(size), mask = Math.ceil(size / 32) * 4 * size, out = Buffer.alloc(40 + size * size * 4 + mask);
  out.writeUInt32LE(40, 0); out.writeInt32LE(size, 4); out.writeInt32LE(size * 2, 8); // the height counts the mask too
  out.writeUInt16LE(1, 12); out.writeUInt16LE(32, 14); out.writeUInt32LE(size * size * 4 + mask, 20);
  for (let y = 0; y < size; y++) for (let x = 0; x < size; x++) {
    const s = (y * size + x) * 4, d = 40 + ((size - 1 - y) * size + x) * 4;
    out[d] = px[s + 2]; out[d + 1] = px[s + 1]; out[d + 2] = px[s]; out[d + 3] = px[s + 3];
  }
  return out;
}
// The exe's icon with every size drawn on its own, so Explorer and Task Manager get the crisp small ones
// instead of the 256 px one shrunk.
export function ico(sizes = [16, 20, 24, 32, 40, 48, 64, 128, 256]) {
  const images = sizes.map(s => (s >= 256 ? png(s) : dib(s)));
  const head = Buffer.alloc(6 + 16 * sizes.length);
  head.writeUInt16LE(1, 2); head.writeUInt16LE(sizes.length, 4); // type 1 = icon
  let offset = head.length;
  sizes.forEach((s, i) => {
    const e = 6 + 16 * i;
    head[e] = head[e + 1] = s % 256; // 0 stands for 256
    head.writeUInt16LE(1, e + 4); head.writeUInt16LE(32, e + 6); // planes, bits per pixel
    head.writeUInt32LE(images[i].length, e + 8); head.writeUInt32LE(offset, e + 12);
    offset += images[i].length;
  });
  return Buffer.concat([head, ...images]);
}
