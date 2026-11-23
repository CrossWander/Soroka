// App and tray icon, drawn in code: the magpie (soroka) of the empty-state picture perched on Teams' purple chat bubble
// with a typing "…", optionally a red badge like Teams'. Pure Node (no Electron), so the build can render build/icon.png with it too.
import zlib from 'node:zlib';

type RGBA = [number, number, number, number];
type Shape = (x: number, y: number) => boolean;
type Paint = RGBA | ((x: number, y: number) => RGBA);
type Pt = [number, number];
const PURPLE: RGBA = [91, 95, 199, 255], GREY: RGBA = [138, 136, 134, 255], WHITE: RGBA = [255, 255, 255, 255], RED: RGBA = [209, 52, 56, 255];
const hex = (h: string, a = 1): RGBA => [1, 3, 5].map(i => parseInt(h.slice(i, i + 2), 16)).concat(255 * a) as RGBA;

const rrect = (x0: number, y0: number, x1: number, y1: number, r: number): Shape => (x, y) => {
  const dx = Math.max(x0 + r - x, 0, x - (x1 - r)), dy = Math.max(y0 + r - y, 0, y - (y1 - r));
  return x >= x0 && x <= x1 && y >= y0 && y <= y1 && dx * dx + dy * dy <= r * r;
};
const circle = (cx: number, cy: number, r: number): Shape => (x, y) => (x - cx) ** 2 + (y - cy) ** 2 <= r * r;
const ellipse = (cx: number, cy: number, rx: number, ry: number, deg = 0): Shape => {
  const c = Math.cos((deg * Math.PI) / 180), s = Math.sin((deg * Math.PI) / 180);
  return (x, y) => { const dx = x - cx, dy = y - cy; return ((dx * c + dy * s) / rx) ** 2 + ((dy * c - dx * s) / ry) ** 2 <= 1; };
};
const any = (...shapes: Shape[]): Shape => (x, y) => shapes.some(s => s(x, y));
const tri = (ax: number, ay: number, bx: number, by: number, cx: number, cy: number): Shape => (x, y) => {
  const side = (px: number, py: number, qx: number, qy: number) => (x - qx) * (py - qy) - (px - qx) * (y - qy);
  const d = [side(ax, ay, bx, by), side(bx, by, cx, cy), side(cx, cy, ax, ay)];
  return !(d.some(v => v < 0) && d.some(v => v > 0));
};
// SVG paths (M, L, C, Z) of the empty-state picture's magpie (src/main.ts, pickPicture), in that picture's coordinates,
// as polylines on the 32×32 design grid: on the icon the magpie is 1.25× its picture size, feet on the bubble's top edge.
const at = (x: number, y: number): Pt => [(x * 1.25 - 22) / 8, (y * 1.25 - 22) / 8];
const lw = (w: number) => (w * 1.25) / 8; // a picture line width on the grid
function polylines(d: string): Pt[][] {
  const t = d.match(/[MLCZ]|-?[\d.]+/g) ?? [], out: Pt[][] = [];
  let i = 0, cmd = '', p: Pt = [0, 0];
  const n = () => +t[i++];
  while (i < t.length) {
    if (/[MLCZ]/.test(t[i])) cmd = t[i++];
    if (cmd === 'M') { p = at(n(), n()); out.push([p]); cmd = 'L'; }
    else if (cmd === 'L') { p = at(n(), n()); out[out.length - 1].push(p); }
    else if (cmd === 'C') { // a cubic Bézier as 8 straight pieces
      const s = p, a = at(n(), n()), b = at(n(), n()), c = at(n(), n());
      for (let j = 1; j <= 8; j++) {
        const u = j / 8, v = 1 - u, m = [v * v * v, 3 * v * v * u, 3 * v * u * u, u * u * u];
        out[out.length - 1].push([0, 1].map(k => m[0] * s[k] + m[1] * a[k] + m[2] * b[k] + m[3] * c[k]) as Pt);
      }
      p = c;
    }
  }
  return out;
}
const box = (pts: Pt[], pad = 0) => {
  const xs = pts.map(q => q[0]), ys = pts.map(q => q[1]);
  return { x0: Math.min(...xs) - pad, x1: Math.max(...xs) + pad, y0: Math.min(...ys) - pad, y1: Math.max(...ys) + pad };
};
// Plain loops over flat arrays: these run 16 times per pixel, and the window icons are drawn when the app starts.
const fill = (d: string): Shape => { // even-odd, as the paths have no holes
  const poly = polylines(d)[0], b = box(poly), xs = Float64Array.from(poly, q => q[0]), ys = Float64Array.from(poly, q => q[1]);
  return (x, y) => {
    if (x < b.x0 || x > b.x1 || y < b.y0 || y > b.y1) return false;
    let inside = false;
    for (let i = 0, j = xs.length - 1; i < xs.length; j = i++)
      if (ys[i] > y !== ys[j] > y && x < ((xs[j] - xs[i]) * (y - ys[i])) / (ys[j] - ys[i]) + xs[i]) inside = !inside;
    return inside;
  };
};
const stroke = (lines: Pt[][], w: number): Shape => { // round caps and joins
  const h = w / 2, b = box(lines.flat(), h); // per segment: its ends, then its box grown by h
  const s = Float64Array.from(lines.flatMap(l => l.slice(1).flatMap(([bx, by], i) => {
    const [ax, ay] = l[i];
    return [ax, ay, bx, by, Math.min(ax, bx) - h, Math.max(ax, bx) + h, Math.min(ay, by) - h, Math.max(ay, by) + h];
  })));
  return (x, y) => {
    if (x < b.x0 || x > b.x1 || y < b.y0 || y > b.y1) return false;
    for (let i = 0; i < s.length; i += 8) {
      if (x < s[i + 4] || x > s[i + 5] || y < s[i + 6] || y > s[i + 7]) continue;
      const ax = s[i], ay = s[i + 1], dx = s[i + 2] - ax, dy = s[i + 3] - ay;
      const t = Math.max(0, Math.min(1, ((x - ax) * dx + (y - ay) * dy) / (dx * dx + dy * dy || 1)));
      if ((ax + t * dx - x) ** 2 + (ay + t * dy - y) ** 2 <= h * h) return true;
    }
    return false;
  };
};
// SVG's linearGradient over a path's bounding box, from its top left towards (gx, gy) of the box
const gradient = (d: string, gx: number, gy: number, stops: [number, string][]): Paint => {
  const b = box(polylines(d)[0]), cs = stops.map(([o, c]) => [o, hex(c)] as const), len = gx * gx + gy * gy;
  return (x, y) => {
    const t = Math.max(0, Math.min(1, (((x - b.x0) / (b.x1 - b.x0)) * gx + ((y - b.y0) / (b.y1 - b.y0)) * gy) / len));
    let i = 1;
    while (i < cs.length - 1 && cs[i][0] < t) i++;
    const [o0, c0] = cs[i - 1], [o1, c1] = cs[i], u = Math.max(0, Math.min(1, (t - o0) / (o1 - o0 || 1)));
    return c0.map((v, k) => v + (c1[k] - v) * u) as RGBA;
  };
};

const TAIL = 'M116 80 C142 92 176 110 216 126 C219 128 217 132 212 131 C174 121 138 107 112 95 Z';
const BODY = 'M44 37 C45 26 61 23 68 30 C75 37 85 45 99 51 C111 56 121 66 127 76 L121 88 C105 98 82 100 66 94 C52 88 44 76 43 62 C42 55 42 49 44 45 Z';
const WING = 'M62 58 C78 51 98 56 114 66 C128 75 138 85 145 96 C130 93 110 87 95 81 C80 75 68 68 62 58 Z';
const BEAK = 'M45 39.5 C39 40 31 43 23.5 47 C31 48.8 39 50.2 46 50.6 Z';
const BELLY = 'M49 74 C56 88 76 96 100 94 C110 92 116 89 119 85 C104 81 84 74 66 66 C58 63 51 67 49 74 Z';
const SCAPULAR = 'M64 58 C76 53 92 55 105 62 C97 67 84 69 73 66 C68 64 65 61 64 58 Z';
const OUTLINE = [TAIL, BODY, WING, BEAK].map(d => { const l = polylines(d)[0]; return [...l, l[0]]; });
const MAGPIE: [Shape, Paint][] = [ // bottom to top
  [fill(TAIL), gradient(TAIL, 1, 0.6, [[0, '#164a44'], [0.45, '#2a7f97'], [0.8, '#3d3590'], [1, '#1e1c3a']])], // iridescent
  [stroke(polylines('M88 96 L86 112 M98 97 L97 112 M80 113 L92 112 M92 113 L102 112'), lw(2.4)), hex('#16181d')], // legs
  [fill(BODY), gradient(BODY, 0, 1, [[0, '#2a2e37'], [1, '#0c0d10']])],
  [ellipse(...at(58, 32), lw(8), lw(3.5), -18), hex('#ffffff', 0.16)], // the sheen on its head
  [fill(BELLY), hex('#f4f4f2')],
  [fill(WING), gradient(WING, 1, 1, [[0, '#2b4680'], [0.6, '#1f6a86'], [1, '#14232e']])],
  [fill(SCAPULAR), hex('#f4f4f2')],
  [stroke(polylines('M100 72 C114 78 128 86 140 94 M96 77 C110 83 122 89 134 94'), lw(0.9)), hex('#7fc4e6', 0.6)], // wing sheen
  [fill(BEAK), hex('#0b0c0e')],
  [circle(...at(55.5, 40), lw(2.6)), hex('#050506')], // eye
  [circle(...at(54.6, 39.1), lw(0.9)), hex('#ffffff', 0.9)],
];
const DETAIL = new Set([1, 3, 7, 10].map(i => MAGPIE[i])); // legs, sheens, the eye's glint: lost below 48 px, they only blur it
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
  const k = size / 32;
  // a light rim parts the black magpie from a dark taskbar (lost on a light one): the picture's, but at least 1 px outside the
  // bird; its inner half is under the magpie's own opaque layers
  const rim = stroke(OUTLINE, Math.max(lw(3.2), 2 / k));
  const layers: [Shape, Paint | null][] = [ // design grid 32×32, later layers on top; null cuts a hole
    [any(rrect(2, 14.75, 30, 29.75, 5), tri(11.5, 29.75, 5.75, 31.75, 7, 28)), grey ? GREY : PURPLE], // the bubble, its tail down left
    [any(circle(11.5, 23, 1.75), circle(16.5, 23, 1.75), circle(21.5, 23, 1.75)), WHITE], // typing…
    [rim, hex('#dfe2ec', 0.75)],
    ...MAGPIE.filter(l => size >= 48 || !DETAIL.has(l)), // it faces left, so a badge in the bottom right corner covers its tail
  ];
  if (badge) { // like Teams: a pill in the bottom right corner, cut out of the tile; the digits on whole pixels stay sharp
    const c = Math.max(1, Math.floor(size / 16)), h = 9 * c, w = Math.max(h, (badge.length * 4 - 1) * c + 4 * c);
    const x0 = (size - w) / k, y0 = (size - h) / k, g = c / k;
    layers.push([rrect(x0 - g, y0 - g, 32 + g, 32 + g, h / 2 / k + g), null], [rrect(x0, y0, 32, 32, h / 2 / k), RED],
      [label(badge, x0 + w / 2 / k, y0 + h / 2 / k, g), WHITE]);
  }

  const N = 4, px = Buffer.alloc(size * size * 4);
  for (let y = 0; y < size; y++) for (let x = 0; x < size; x++) {
    let r = 0, g = 0, b = 0, cover = 0;
    for (let sy = 0; sy < N; sy++) for (let sx = 0; sx < N; sx++) { // 4×4 supersampling for smooth edges
      const u = (x + (sx + 0.5) / N) / k, v = (y + (sy + 0.5) / N) / k;
      let left = 1; // what the layers above let through: a translucent layer shows the ones below it
      for (let i = layers.length - 1; i >= 0 && left > 0; i--) {
        const [shape, paint] = layers[i];
        if (!shape(u, v)) continue;
        if (!paint) break;
        const c = typeof paint === 'function' ? paint(u, v) : paint, w = (left * c[3]) / 255;
        r += c[0] * w; g += c[1] * w; b += c[2] * w; cover += w; left -= w;
      }
    }
    const i = (y * size + x) * 4;
    if (cover) { px[i] = r / cover; px[i + 1] = g / cover; px[i + 2] = b / cover; px[i + 3] = (255 * cover) / (N * N); }
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
