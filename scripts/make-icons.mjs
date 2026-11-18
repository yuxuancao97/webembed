// Writes the add-in icons (PNG) with no dependencies: an indigo rounded
// square with a white "screen" frame and a play triangle.
import fs from "node:fs";
import path from "node:path";
import zlib from "node:zlib";
import { fileURLToPath } from "node:url";

const out = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "public", "assets");

const crcTable = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
const crc32 = (buf) => {
  let c = 0xffffffff;
  for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};
const chunk = (type, data) => {
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
};

function icon(size) {
  const S = 4; // supersampling for smooth edges
  const px = Buffer.alloc(size * size * 4);
  const accent = [79, 70, 229];
  for (let y = 0; y < size; y++) for (let x = 0; x < size; x++) {
    let bg = 0, white = 0;
    for (let sy = 0; sy < S; sy++) for (let sx = 0; sx < S; sx++) {
      const u = (x + (sx + 0.5) / S) / size, v = (y + (sy + 0.5) / S) / size;
      // rounded square
      const r = 0.2, dx = Math.max(Math.abs(u - 0.5) - (0.5 - r), 0), dy = Math.max(Math.abs(v - 0.5) - (0.5 - r), 0);
      if (dx * dx + dy * dy > r * r) continue;
      bg++;
      // frame
      const inOuter = u > 0.18 && u < 0.82 && v > 0.24 && v < 0.76;
      const inInner = u > 0.25 && u < 0.75 && v > 0.31 && v < 0.69;
      // play triangle
      const tri = u > 0.43 && u < 0.6 && Math.abs(v - 0.5) < (0.6 - u) * 0.62;
      if ((inOuter && !inInner) || tri) white++;
    }
    const a = bg / (S * S), w = bg ? white / bg : 0, i = (y * size + x) * 4;
    for (let k = 0; k < 3; k++) px[i + k] = Math.round(accent[k] * (1 - w) + 255 * w);
    px[i + 3] = Math.round(a * 255);
  }
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) px.copy(raw, y * (size * 4 + 1) + 1, y * size * 4, (y + 1) * size * 4);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0); ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; ihdr[9] = 6; // 8-bit RGBA
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr), chunk("IDAT", zlib.deflateSync(raw)), chunk("IEND", Buffer.alloc(0)),
  ]);
}

fs.mkdirSync(out, { recursive: true });
for (const s of [16, 32, 64, 80]) fs.writeFileSync(path.join(out, `icon-${s}.png`), icon(s));
console.log("Icons written to", out);
