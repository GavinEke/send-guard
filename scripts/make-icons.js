const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

const root = path.resolve(__dirname, "..");
const assetsDir = path.join(root, "assets");
fs.mkdirSync(assetsDir, { recursive: true });

const sizes = [16, 32, 64, 80, 128];
for (const size of sizes) {
  fs.writeFileSync(path.join(assetsDir, `icon-${size}.png`), makePng(size));
}

console.log(`Generated ${sizes.length} icon files.`);

function makePng(size) {
  const rows = [];
  for (let y = 0; y < size; y += 1) {
    const row = Buffer.alloc(1 + size * 4);
    row[0] = 0;
    for (let x = 0; x < size; x += 1) {
      const offset = 1 + x * 4;
      const inBadge = x > size * 0.2 && x < size * 0.8 && y > size * 0.2 && y < size * 0.8;
      const inCheck =
        inBadge &&
        Math.abs(y - (size * 0.62 - x * 0.34)) < size * 0.08 &&
        x < size * 0.52;
      const inCheckTail =
        inBadge &&
        Math.abs(y - (x * 0.42 + size * 0.18)) < size * 0.08 &&
        x >= size * 0.45;

      const color = inCheck || inCheckTail
        ? [255, 255, 255, 255]
        : inBadge
          ? [24, 90, 189, 255]
          : [246, 247, 249, 0];

      row[offset] = color[0];
      row[offset + 1] = color[1];
      row[offset + 2] = color[2];
      row[offset + 3] = color[3];
    }
    rows.push(row);
  }

  return Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    chunk("IHDR", Buffer.concat([uint32(size), uint32(size), Buffer.from([8, 6, 0, 0, 0])])),
    chunk("IDAT", zlib.deflateSync(Buffer.concat(rows))),
    chunk("IEND", Buffer.alloc(0))
  ]);
}

function chunk(type, data) {
  const typeBuffer = Buffer.from(type, "ascii");
  return Buffer.concat([uint32(data.length), typeBuffer, data, uint32(crc32(Buffer.concat([typeBuffer, data])))]);
}

function uint32(value) {
  const buffer = Buffer.alloc(4);
  buffer.writeUInt32BE(value >>> 0);
  return buffer;
}

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = crc & 1 ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1;
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}
