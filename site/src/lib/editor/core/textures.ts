// PowerPoint's Texture gallery: the twenty-four textures in gallery order.
// Microsoft's bitmaps are not redistributable, so each tile here is drawn
// procedurally to evoke the named material. Every generator is a function of
// the pixel position that repeats every TEXTURE_SIZE pixels (periodic noise
// lattices, sine terms with a whole number of cycles per tile, wrapped
// distances), so the tiles wrap seamlessly.

/** PowerPoint's default texture tile: 128 × 128 pixels. */
export const TEXTURE_SIZE = 128;
// PowerPoint's texture media is ~144 DPI, so a tile spans 64 pt (pHYs is per metre).
const TEXTURE_DPI = 144;
const INCHES_PER_METRE = 39.3701;
// tEXt keyword naming the texture, so the gallery can mark the current fill.
const TEXTURE_KEYWORD = 'office-kit texture';

export const TEXTURES = [
  { id: 'papyrus', name: 'Papyrus' },
  { id: 'canvas', name: 'Canvas' },
  { id: 'denim', name: 'Denim' },
  { id: 'woven-mat', name: 'Woven mat' },
  { id: 'water-droplets', name: 'Water droplets' },
  { id: 'paper-bag', name: 'Paper bag' },
  { id: 'fish-fossil', name: 'Fish fossil' },
  { id: 'sand', name: 'Sand' },
  { id: 'green-marble', name: 'Green marble' },
  { id: 'white-marble', name: 'White marble' },
  { id: 'brown-marble', name: 'Brown marble' },
  { id: 'granite', name: 'Granite' },
  { id: 'newsprint', name: 'Newsprint' },
  { id: 'recycled-paper', name: 'Recycled paper' },
  { id: 'parchment', name: 'Parchment' },
  { id: 'stationery', name: 'Stationery' },
  { id: 'blue-tissue-paper', name: 'Blue tissue paper' },
  { id: 'pink-tissue-paper', name: 'Pink tissue paper' },
  { id: 'purple-mesh', name: 'Purple mesh' },
  { id: 'bouquet', name: 'Bouquet' },
  { id: 'cork', name: 'Cork' },
  { id: 'walnut', name: 'Walnut' },
  { id: 'oak', name: 'Oak' },
  { id: 'medium-wood', name: 'Medium wood' },
] as const;

export type TextureId = (typeof TEXTURES)[number]['id'];

/**
 * What PowerPoint inserts when Picture or texture fill is chosen for a shape or
 * background with no picture to restore: Papyrus, the gallery's first texture
 * (see test/fixtures/native/texture-capture.md).
 */
export const DEFAULT_TEXTURE: TextureId = 'papyrus';
type Rgb = readonly [number, number, number];
type Field = (x: number, y: number) => Rgb;
type Random = () => number;

const N = TEXTURE_SIZE;
const TAU = Math.PI * 2;

function random(seed: number): Random {
  // mulberry32: small, fast and identical on every engine.
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let value = state;
    value = Math.imul(value ^ (value >>> 15), value | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
  };
}

const hex = (value: string): Rgb => {
  const channel = (at: number) => parseInt(value.slice(at, at + 2), 16);
  return [channel(1), channel(3), channel(5)];
};
const clamp01 = (value: number) => Math.min(1, Math.max(0, value));
const mix = (a: Rgb, b: Rgb, amount: number): Rgb => {
  const t = clamp01(amount);
  return [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t, a[2] + (b[2] - a[2]) * t];
};
const shade = (color: Rgb, factor: number): Rgb => [
  color[0] * factor,
  color[1] * factor,
  color[2] * factor,
];
const smooth = (t: number) => t * t * (3 - 2 * t);
const wrap = (value: number, period: number) => ((value % period) + period) % period;
/** Shortest distance across the tile's wrap. */
const wrapped = (delta: number) => {
  const d = Math.abs(delta) % N;
  return Math.min(d, N - d);
};

/** Value noise with `cellsX × cellsY` lattice cells per tile; periodic over the tile. */
function noise(rng: Random, cellsX: number, cellsY = cellsX): (x: number, y: number) => number {
  const lattice = Float64Array.from({ length: cellsX * cellsY }, rng);
  const sx = cellsX / N,
    sy = cellsY / N;
  return (x, y) => {
    const gx = x * sx,
      gy = y * sy;
    const x0 = Math.floor(gx),
      y0 = Math.floor(gy);
    const tx = smooth(gx - x0),
      ty = smooth(gy - y0);
    const at = (i: number, j: number) => lattice[wrap(j, cellsY) * cellsX + wrap(i, cellsX)]!;
    const top = at(x0, y0) + (at(x0 + 1, y0) - at(x0, y0)) * tx;
    const bottom = at(x0, y0 + 1) + (at(x0 + 1, y0 + 1) - at(x0, y0 + 1)) * tx;
    return top + (bottom - top) * ty;
  };
}

/** Fractal sum of octaves, each doubling the lattice; normalised to 0..1. */
function fbm(rng: Random, cellsX: number, cellsY: number, octaves: number, persistence = 0.5) {
  const layers = Array.from({ length: octaves }, (_, octave) =>
    noise(rng, cellsX << octave, cellsY << octave),
  );
  const total = layers.reduce((sum, _, octave) => sum + persistence ** octave, 0);
  return (x: number, y: number) =>
    layers.reduce((sum, layer, octave) => sum + layer(x, y) * persistence ** octave, 0) / total;
}

/** Independent per-pixel noise (already periodic: one value per pixel of the tile). */
function grain(rng: Random): (x: number, y: number) => number {
  const values = Float64Array.from({ length: N * N }, rng);
  return (x, y) => values[wrap(y, N) * N + wrap(x, N)]!;
}

interface Spot {
  x: number;
  y: number;
  radius: number;
  value: number;
}
function spots(rng: Random, count: number, minRadius: number, maxRadius: number): Spot[] {
  return Array.from({ length: count }, () => ({
    x: rng() * N,
    y: rng() * N,
    radius: minRadius + rng() * (maxRadius - minRadius),
    value: rng(),
  }));
}
/** The first spot covering (x, y), with the offset to its centre across the wrap. */
function spotAt(
  list: readonly Spot[],
  x: number,
  y: number,
): { spot: Spot; dx: number; dy: number } | null {
  for (const spot of list) {
    const dx = wrapped(x - spot.x),
      dy = wrapped(y - spot.y);
    if (dx * dx + dy * dy <= spot.radius * spot.radius) {
      return {
        spot,
        dx: wrap(x - spot.x + N / 2, N) - N / 2,
        dy: wrap(y - spot.y + N / 2, N) - N / 2,
      };
    }
  }
  return null;
}

function marble(seed: number, base: Rgb, mid: Rgb, vein: Rgb): Field {
  const rng = random(seed);
  const turbulence = fbm(rng, 2, 2, 5);
  const cloud = fbm(rng, 4, 4, 3);
  return (x, y) => {
    // Two cycles across, one down: a diagonal grain that still repeats per tile.
    const phase = (TAU * (2 * x + y)) / N + 9 * turbulence(x, y);
    const veins = (1 - Math.abs(Math.sin(phase))) ** 10;
    return mix(mix(base, mid, cloud(x, y) * 1.4 - 0.2), vein, veins);
  };
}

function wood(seed: number, rings: number, dark: Rgb, light: Rgb, sharpness: number): Field {
  const rng = random(seed);
  // Few cells along x, many along y: grain runs horizontally.
  const warp = fbm(rng, 1, 4, 4);
  const fibre = noise(rng, 8, 64);
  return (x, y) => {
    const ring = 0.5 + 0.5 * Math.sin((TAU * rings * y) / N + 7 * warp(x, y));
    return shade(mix(dark, light, ring ** sharpness), 0.9 + 0.2 * fibre(x, y));
  };
}

function tissue(seed: number, color: Rgb): Field {
  const rng = random(seed);
  const crease = fbm(rng, 4, 4, 4);
  const fine = grain(rng);
  return (x, y) => {
    const ridge = 1 - Math.abs(2 * crease(x, y) - 1);
    return shade(color, 0.86 + 0.2 * ridge ** 3 + 0.03 * fine(x, y));
  };
}

const GENERATORS: Record<TextureId, (seed: number) => Field> = {
  papyrus: (seed) => {
    const rng = random(seed);
    const across = fbm(rng, 2, 32, 3),
      down = fbm(rng, 32, 2, 3),
      fine = grain(rng);
    const dark = hex('#a88a52'),
      light = hex('#ead9a6');
    return (x, y) =>
      mix(dark, light, 0.55 * across(x, y) + 0.35 * down(x, y) + 0.1 * fine(x, y) + 0.15);
  },
  canvas: (seed) => {
    const rng = random(seed);
    const fine = grain(rng),
      cloud = fbm(rng, 4, 4, 2);
    const color = hex('#d2c6a8');
    return (x, y) => {
      // 4-pixel threads crossing: 32 whole cycles per tile.
      const weave = Math.cos((TAU * x) / 4) * Math.cos((TAU * y) / 4);
      return shade(color, 0.92 + 0.07 * weave + 0.06 * fine(x, y) + 0.05 * cloud(x, y));
    };
  },
  denim: (seed) => {
    const rng = random(seed);
    const fine = grain(rng),
      cloud = fbm(rng, 4, 4, 3);
    const navy = hex('#2b4570'),
      faded = hex('#7690b8'),
      thread = hex('#dfe6f0');
    return (x, y) => {
      const twill = Math.sin((TAU * (x + y)) / 4);
      const base = mix(navy, faded, 0.3 + 0.2 * twill + 0.35 * cloud(x, y) - 0.15);
      return fine(x, y) > 0.93 ? mix(base, thread, 0.6) : base;
    };
  },
  'woven-mat': (seed) => {
    const rng = random(seed);
    const fine = grain(rng);
    const shadow = hex('#80602f'),
      straw = hex('#d9bd80');
    const cell = 16,
      strand = cell / 3;
    return (x, y) => {
      // Eight cells per tile, so the alternating strand direction wraps.
      const horizontal = (Math.floor(x / cell) + Math.floor(y / cell)) % 2 === 0;
      const across = horizontal ? y : x,
        along = horizontal ? x : y;
      const profile = Math.sin((Math.PI * wrap(across, strand)) / strand);
      const edge = Math.sin((Math.PI * wrap(along, cell)) / cell);
      return mix(shadow, straw, 0.25 + 0.6 * profile * (0.6 + 0.4 * edge) + 0.15 * fine(x, y));
    };
  },
  'water-droplets': (seed) => {
    const rng = random(seed);
    const cloud = fbm(rng, 2, 2, 3);
    const drops = spots(rng, 46, 2, 7);
    const deep = hex('#5f7f9c'),
      pale = hex('#a9bfd1'),
      highlight = hex('#f4f8fb');
    return (x, y) => {
      const base = mix(deep, pale, cloud(x, y));
      const hit = spotAt(drops, x, y);
      if (!hit) return base;
      // Light from the upper left on a hemisphere.
      const nx = hit.dx / hit.spot.radius,
        ny = hit.dy / hit.spot.radius;
      const nz = Math.sqrt(Math.max(0, 1 - nx * nx - ny * ny));
      const light = clamp01(-0.55 * nx - 0.55 * ny + 0.62 * nz);
      return mix(shade(base, 0.75 + 0.45 * light), highlight, light ** 12);
    };
  },
  'paper-bag': (seed) => {
    const rng = random(seed);
    const fibre = fbm(rng, 16, 2, 3),
      fine = grain(rng),
      cloud = fbm(rng, 4, 4, 2);
    const kraft = hex('#b8936a');
    return (x, y) =>
      shade(kraft, 0.9 + 0.08 * fibre(x, y) + 0.06 * fine(x, y) + 0.06 * cloud(x, y));
  },
  'fish-fossil': (seed) => {
    const rng = random(seed);
    const stone = fbm(rng, 4, 4, 4),
      fine = grain(rng);
    const light = hex('#d8c6a0'),
      dark = hex('#9c8458'),
      imprint = hex('#6f5a3a');
    return (x, y) => {
      const base = mix(dark, light, stone(x, y) * 1.3 - 0.1 + 0.08 * fine(x, y));
      // A fish imprint centred in the tile, clear of the edges.
      const bx = (x - 60) / 38,
        by = (y - 64) / 14;
      const body = bx * bx + by * by;
      const tail = x > 94 && x < 112 && Math.abs(y - 64) < (x - 94) * 0.8;
      if (body < 1) {
        const rib = Math.abs(Math.sin((TAU * x) / 6)) > 0.75 || Math.abs(y - 64) < 1;
        return mix(base, imprint, rib ? 0.75 : body > 0.8 ? 0.5 : 0.15);
      }
      return tail ? mix(base, imprint, 0.35) : base;
    };
  },
  sand: (seed) => {
    const rng = random(seed);
    const fine = grain(rng),
      cloud = fbm(rng, 4, 4, 3);
    const color = hex('#d6c095'),
      speck = hex('#6e5b3d');
    return (x, y) => {
      const value = fine(x, y);
      const base = shade(color, 0.85 + 0.2 * value + 0.08 * cloud(x, y));
      return value > 0.97 ? mix(base, speck, 0.6) : base;
    };
  },
  'green-marble': (seed) => marble(seed, hex('#244a38'), hex('#4f8566'), hex('#bcd8c2')),
  'white-marble': (seed) => marble(seed, hex('#f0efeb'), hex('#d6d3cc'), hex('#8c8984')),
  'brown-marble': (seed) => marble(seed, hex('#5e3520'), hex('#8e5735'), hex('#dbb58d')),
  granite: (seed) => {
    const rng = random(seed);
    const flecks = noise(rng, 64, 64),
      cloud = fbm(rng, 4, 4, 3),
      fine = grain(rng);
    const gray = hex('#8d8883'),
      black = hex('#35312f'),
      white = hex('#ddd7d1'),
      pink = hex('#b48d82');
    return (x, y) => {
      const value = flecks(x, y) + 0.15 * (fine(x, y) - 0.5);
      const base = shade(gray, 0.9 + 0.2 * cloud(x, y));
      if (value > 0.72) return black;
      if (value < 0.2) return white;
      return value > 0.62 ? pink : base;
    };
  },
  newsprint: (seed) => {
    const rng = random(seed);
    const fine = grain(rng),
      cloud = fbm(rng, 4, 4, 2);
    // One "word" switch per 4-pixel block: 16 lines of 32 blocks.
    const words = Array.from({ length: 16 * 32 }, () => rng() > 0.25);
    const paper = hex('#dedbd3'),
      ink = hex('#77746d');
    return (x, y) => {
      const base = shade(paper, 0.94 + 0.06 * cloud(x, y) + 0.04 * fine(x, y));
      const row = wrap(y, 8);
      const inked =
        row >= 2 &&
        row <= 5 &&
        words[Math.floor(y / 8) * 32 + Math.floor(x / 4)] &&
        fine(x, y) > 0.35;
      return inked ? mix(base, ink, 0.55) : base;
    };
  },
  'recycled-paper': (seed) => {
    const rng = random(seed);
    const fine = grain(rng),
      cloud = fbm(rng, 4, 4, 3);
    const flecks = spots(rng, 70, 0.6, 1.6);
    const paper = hex('#e3dccb');
    const colors = [hex('#8b7355'), hex('#7d8794'), hex('#6f86a6'), hex('#a08f78')];
    return (x, y) => {
      const base = shade(paper, 0.93 + 0.05 * cloud(x, y) + 0.04 * fine(x, y));
      const hit = spotAt(flecks, x, y);
      return hit ? mix(base, colors[Math.floor(hit.spot.value * colors.length)]!, 0.7) : base;
    };
  },
  parchment: (seed) => {
    const rng = random(seed);
    const blotch = fbm(rng, 2, 2, 5, 0.55),
      fine = grain(rng);
    const light = hex('#efdfb4'),
      dark = hex('#c19c62');
    return (x, y) => mix(light, dark, (blotch(x, y) - 0.3) * 1.6 + 0.05 * fine(x, y));
  },
  stationery: (seed) => {
    const rng = random(seed);
    const fibre = fbm(rng, 8, 8, 4),
      fine = grain(rng);
    const paper = hex('#f2eacf');
    return (x, y) => shade(paper, 0.95 + 0.05 * fibre(x, y) + 0.025 * fine(x, y));
  },
  'blue-tissue-paper': (seed) => tissue(seed, hex('#a6c4e2')),
  'pink-tissue-paper': (seed) => tissue(seed, hex('#f0bccd')),
  'purple-mesh': (seed) => {
    const rng = random(seed);
    const cloud = fbm(rng, 4, 4, 3);
    const hole = hex('#3b2353'),
      thread = hex('#a582c8');
    return (x, y) => {
      // 8-pixel mesh: sixteen cells per tile.
      const tx = Math.cos((TAU * x) / 8),
        ty = Math.cos((TAU * y) / 8);
      const strand = Math.max(tx, ty);
      return shade(mix(hole, thread, (strand - 0.2) * 1.4), 0.85 + 0.25 * cloud(x, y));
    };
  },
  bouquet: (seed) => {
    const rng = random(seed);
    const cloud = fbm(rng, 4, 4, 3);
    const flowers = spots(rng, 18, 7, 12);
    const leaves = spots(rng, 26, 4, 7);
    const ground = hex('#efe2dc'),
      leaf = hex('#6f9a5a'),
      centre = hex('#f2cf4a');
    const petals = [hex('#e27aa4'), hex('#b07ad0'), hex('#f3a3b8'), hex('#f7f2f4'), hex('#d9566f')];
    return (x, y) => {
      const flower = spotAt(flowers, x, y);
      if (flower) {
        const r = Math.hypot(flower.dx, flower.dy) / flower.spot.radius;
        const angle = Math.atan2(flower.dy, flower.dx) + flower.spot.value * TAU;
        if (r < 0.22) return centre;
        if (r < 0.55 + 0.45 * Math.abs(Math.cos(2.5 * angle))) {
          return shade(petals[Math.floor(flower.spot.value * petals.length)]!, 1 - 0.25 * r);
        }
      }
      const green = spotAt(leaves, x, y);
      if (green) return shade(leaf, 0.8 + 0.3 * green.spot.value);
      return shade(ground, 0.95 + 0.06 * cloud(x, y));
    };
  },
  cork: (seed) => {
    const rng = random(seed);
    const cells = spots(rng, 300, 0, 0);
    const fine = grain(rng),
      cloud = fbm(rng, 4, 4, 2);
    const cork = hex('#c0905a'),
      dark = hex('#6b4422');
    return (x, y) => {
      let nearest = Infinity,
        second = Infinity;
      for (const cell of cells) {
        const dx = wrapped(x - cell.x),
          dy = wrapped(y - cell.y);
        const distance = dx * dx + dy * dy;
        if (distance < nearest) {
          second = nearest;
          nearest = distance;
        } else if (distance < second) second = distance;
      }
      // Granule boundaries darken; pores are the darkest grain values.
      const boundary = clamp01(1 - (Math.sqrt(second) - Math.sqrt(nearest)) / 1.5);
      const base = shade(cork, 0.88 + 0.18 * cloud(x, y) + 0.08 * fine(x, y));
      return mix(base, dark, 0.45 * boundary + (fine(x, y) > 0.96 ? 0.5 : 0));
    };
  },
  walnut: (seed) => wood(seed, 9, hex('#3f2414'), hex('#7b5131'), 1.5),
  oak: (seed) => wood(seed, 6, hex('#8a5a2e'), hex('#d0a46c'), 0.6),
  'medium-wood': (seed) => wood(seed, 12, hex('#6a4020'), hex('#ad7a46'), 1),
};

/** The texture's RGBA pixels, row-major, `TEXTURE_SIZE` square. Deterministic. */
export function texturePixels(id: TextureId): Uint8ClampedArray<ArrayBuffer> {
  const index = TEXTURES.findIndex((texture) => texture.id === id);
  const field = GENERATORS[id](0x9e3779b9 ^ (index * 0x85ebca6b));
  const pixels = new Uint8ClampedArray(N * N * 4);
  for (let y = 0; y < N; y++) {
    for (let x = 0; x < N; x++) {
      const [r, g, b] = field(x, y);
      const at = (y * N + x) * 4;
      pixels[at] = r;
      pixels[at + 1] = g;
      pixels[at + 2] = b;
      pixels[at + 3] = 255;
    }
  }
  return pixels;
}

const CRC_TABLE = Uint32Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = CRC_TABLE[(crc ^ byte) & 0xff]! ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}
function chunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.length);
  out.set(new TextEncoder().encode(type), 4);
  out.set(data, 8);
  view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
  return out;
}
const PNG_SIGNATURE = [137, 80, 78, 71, 13, 10, 26, 10];

async function zlib(data: Uint8Array): Promise<Uint8Array> {
  // CompressionStream's "deflate" is the zlib wrapper PNG's IDAT expects.
  const stream = new Blob([data as BlobPart])
    .stream()
    .pipeThrough(new CompressionStream('deflate'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** Encodes the texture as an RGB PNG at ~144 DPI, tagged with its id. */
export async function encodeTexturePng(id: TextureId): Promise<Uint8Array> {
  const pixels = texturePixels(id);
  const raw = new Uint8Array(N * (1 + N * 3));
  for (let y = 0; y < N; y++) {
    // Filter byte 0 (None), then the row's RGB.
    const row = y * (1 + N * 3);
    for (let x = 0; x < N; x++)
      raw.set(pixels.subarray((y * N + x) * 4, (y * N + x) * 4 + 3), row + 1 + x * 3);
  }
  const header = new Uint8Array(13);
  const headerView = new DataView(header.buffer);
  headerView.setUint32(0, N);
  headerView.setUint32(4, N);
  header.set([8, 2, 0, 0, 0], 8); // 8-bit truecolour, deflate, adaptive filtering, no interlace
  const density = new Uint8Array(9);
  const densityView = new DataView(density.buffer);
  const perMetre = Math.round(TEXTURE_DPI * INCHES_PER_METRE);
  densityView.setUint32(0, perMetre);
  densityView.setUint32(4, perMetre);
  density[8] = 1; // unit: metre
  const text = new TextEncoder().encode(`${TEXTURE_KEYWORD}\0${id}`);
  const parts = [
    new Uint8Array(PNG_SIGNATURE),
    chunk('IHDR', header),
    chunk('pHYs', density),
    chunk('tEXt', text),
    chunk('IDAT', await zlib(raw)),
    chunk('IEND', new Uint8Array(0)),
  ];
  const png = new Uint8Array(parts.reduce((length, part) => length + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    png.set(part, offset);
    offset += part.length;
  }
  return png;
}

const encoded = new Map<TextureId, Promise<Uint8Array>>();
/** The texture's PNG, generated on first use and cached for the session. */
export function texturePng(id: TextureId): Promise<Uint8Array> {
  let png = encoded.get(id);
  if (!png) {
    png = encodeTexturePng(id);
    encoded.set(id, png);
  }
  return png;
}

/** The gallery texture a fill image was made from, read from its PNG tag; null for any other image. */
export function textureIdOf(bytes: Uint8Array | null): TextureId | null {
  if (!bytes || bytes.length < 8 || PNG_SIGNATURE.some((value, index) => bytes[index] !== value))
    return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const decoder = new TextDecoder('latin1');
  for (let offset = 8; offset + 8 <= bytes.length; ) {
    const length = view.getUint32(offset);
    const type = decoder.decode(bytes.subarray(offset + 4, offset + 8));
    if (type === 'IDAT' || type === 'IEND') return null;
    if (type === 'tEXt') {
      const [keyword, value] = decoder
        .decode(bytes.subarray(offset + 8, offset + 8 + length))
        .split('\0');
      if (keyword === TEXTURE_KEYWORD)
        return TEXTURES.find((texture) => texture.id === value)?.id ?? null;
    }
    offset += 12 + length;
  }
  return null;
}
