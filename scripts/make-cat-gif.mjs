// Builds assets/cat-walk.gif from assets/cat-sprite.png (a horizontal strip of square walk-cycle frames).
// Dependency-free: decodes the PNG with node:zlib and writes the GIF (LZW) by hand.
// Also writes src/cat-walk-gif.ts, the base64 data URI the empty-state card embeds.
// Usage: node scripts/make-cat-gif.mjs [sprite.png] [out.gif]
import { readFileSync, writeFileSync } from "node:fs";
import { inflateSync } from "node:zlib";

const SPRITE = process.argv[2] ?? "assets/cat-sprite.png";
const OUT = process.argv[3] ?? "assets/cat-walk.gif";
const OUT_TS = "src/cat-walk-gif.ts";

const FRAME = 50; // sprite cell size in px
const SCALE = 3; // nearest-neighbour upscale
const CANVAS_W = 280; // sidebar-friendly width, in output px
const STEP = 3; // source px the cat travels per animation frame
const FRAME_MS = 100;
const PAUSE_MS = 1500; // empty frame at the end so the loop breathes

function decodePng(buf) {
  let pos = 8;
  let width = 0;
  let height = 0;
  const idat = [];
  while (pos < buf.length) {
    const len = buf.readUInt32BE(pos);
    const type = buf.toString("ascii", pos + 4, pos + 8);
    const data = buf.subarray(pos + 8, pos + 8 + len);
    if (type === "IHDR") {
      width = data.readUInt32BE(0);
      height = data.readUInt32BE(4);
      if (data[8] !== 8 || data[9] !== 6 || data[12] !== 0) {
        throw new Error("Sprite must be 8-bit RGBA, non-interlaced PNG");
      }
    } else if (type === "IDAT") {
      idat.push(data);
    }
    pos += 12 + len;
  }
  const raw = inflateSync(Buffer.concat(idat));
  const stride = width * 4;
  const px = Buffer.alloc(height * stride);
  for (let y = 0; y < height; y++) {
    const filter = raw[y * (stride + 1)];
    for (let x = 0; x < stride; x++) {
      const v = raw[y * (stride + 1) + 1 + x];
      const a = x >= 4 ? px[y * stride + x - 4] : 0;
      const b = y ? px[(y - 1) * stride + x] : 0;
      const c = x >= 4 && y ? px[(y - 1) * stride + x - 4] : 0;
      let out;
      if (filter === 0) out = v;
      else if (filter === 1) out = v + a;
      else if (filter === 2) out = v + b;
      else if (filter === 3) out = v + ((a + b) >> 1);
      else {
        const p = a + b - c;
        const pa = Math.abs(p - a);
        const pb = Math.abs(p - b);
        const pc = Math.abs(p - c);
        out = v + (pa <= pb && pa <= pc ? a : pb <= pc ? b : c);
      }
      px[y * stride + x] = out & 255;
    }
  }
  return { width, height, px };
}

function lzw(indices, minCode) {
  const clear = 1 << minCode;
  const eoi = clear + 1;
  const bytes = [];
  let bitBuf = 0;
  let bitCnt = 0;
  let size = minCode + 1;
  const write = (code) => {
    bitBuf |= code << bitCnt;
    bitCnt += size;
    while (bitCnt >= 8) {
      bytes.push(bitBuf & 255);
      bitBuf >>>= 8;
      bitCnt -= 8;
    }
  };
  let dict = new Map();
  let next = eoi + 1;
  const reset = () => {
    dict = new Map();
    next = eoi + 1;
    size = minCode + 1;
  };
  write(clear);
  let cur = indices[0];
  for (let i = 1; i < indices.length; i++) {
    const k = indices[i];
    const key = (cur << 8) | k;
    if (dict.has(key)) {
      cur = dict.get(key);
      continue;
    }
    write(cur);
    if (next < 4096) {
      dict.set(key, next++);
      if (next - 1 === 1 << size && size < 12) size++;
    } else {
      write(clear);
      reset();
    }
    cur = k;
  }
  write(cur);
  write(eoi);
  if (bitCnt > 0) bytes.push(bitBuf & 255);
  return bytes;
}

const { width: sw, height: sh, px } = decodePng(readFileSync(SPRITE));
const frameCount = Math.floor(sw / FRAME);

// Palette: index 0 is transparent, then each distinct opaque colour.
const palette = [[0, 0, 0]];
const paletteIndex = new Map();
const sprite = new Uint8Array(sw * sh); // palette index per source pixel
for (let i = 0; i < sw * sh; i++) {
  if (px[i * 4 + 3] < 128) continue;
  const key = (px[i * 4] << 16) | (px[i * 4 + 1] << 8) | px[i * 4 + 2];
  if (!paletteIndex.has(key)) {
    paletteIndex.set(key, palette.length);
    palette.push([px[i * 4], px[i * 4 + 1], px[i * 4 + 2]]);
  }
  sprite[i] = paletteIndex.get(key);
}
if (palette.length > 256) throw new Error("Too many colours for a GIF palette");

// Crop to the rows the cat actually occupies so the GIF isn't mostly empty.
let top = sh;
let bottom = -1;
let minX = FRAME;
let maxX = -1;
for (let y = 0; y < sh; y++) {
  for (let x = 0; x < sw; x++) {
    if (!sprite[y * sw + x]) continue;
    top = Math.min(top, y);
    bottom = Math.max(bottom, y);
    minX = Math.min(minX, x % FRAME);
    maxX = Math.max(maxX, x % FRAME);
  }
}
const catW = maxX - minX + 1;
const cropH = bottom - top + 1;
const outW = CANVAS_W;
const outH = cropH * SCALE;

function renderFrame(f, offsetSrcX) {
  const out = new Uint8Array(outW * outH);
  if (f === null) return out;
  for (let y = 0; y < cropH; y++) {
    for (let x = 0; x < catW; x++) {
      const idx = sprite[(top + y) * sw + f * FRAME + minX + x];
      if (!idx) continue;
      const ox = (offsetSrcX + x) * SCALE;
      for (let dy = 0; dy < SCALE; dy++) {
        for (let dx = 0; dx < SCALE; dx++) {
          const tx = ox + dx;
          if (tx >= 0 && tx < outW) out[(y * SCALE + dy) * outW + tx] = idx;
        }
      }
    }
  }
  return out;
}

// Walk from just off the left edge to just off the right edge.
const startX = -catW + STEP;
const endX = Math.ceil(outW / SCALE);
const frames = [];
for (let n = 0, x = startX; x < endX; n++, x += STEP) {
  frames.push({ pixels: renderFrame(n % frameCount, x), delay: FRAME_MS });
}
frames.push({ pixels: renderFrame(null, 0), delay: PAUSE_MS });

let bits = 1;
while (1 << bits < palette.length) bits++;
bits = Math.max(bits, 2);
const table = [];
for (let i = 0; i < 1 << bits; i++) table.push(...(palette[i] ?? [0, 0, 0]));

const out = [];
const u16 = (n) => out.push(n & 255, n >> 8);
out.push(...Buffer.from("GIF89a"));
u16(outW);
u16(outH);
out.push(0x80 | ((bits - 1) << 4) | (bits - 1), 0, 0);
out.push(...table);
out.push(0x21, 0xff, 0x0b, ...Buffer.from("NETSCAPE2.0"), 3, 1, 0, 0, 0); // loop forever
for (const frame of frames) {
  // disposal 2 (clear to transparent), transparency on index 0
  out.push(0x21, 0xf9, 4, (2 << 2) | 1);
  u16(Math.round(frame.delay / 10));
  out.push(0, 0);
  out.push(0x2c);
  u16(0);
  u16(0);
  u16(outW);
  u16(outH);
  out.push(0, bits);
  const data = lzw(frame.pixels, bits);
  for (let i = 0; i < data.length; i += 255) {
    const chunk = data.slice(i, i + 255);
    out.push(chunk.length, ...chunk);
  }
  out.push(0);
}
out.push(0x3b);
const gif = Buffer.from(out);
writeFileSync(OUT, gif);
writeFileSync(
  OUT_TS,
  `// Generated by scripts/make-cat-gif.mjs from assets/cat-sprite.png. Do not edit.\nexport const CAT_WALK_GIF_DATA_URI =\n  "data:image/gif;base64,${gif.toString("base64")}";\n`,
);
console.log(`${OUT}: ${outW}x${outH}, ${frames.length} frames, ${out.length} bytes`);
