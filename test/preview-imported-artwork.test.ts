import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideImage,
  addSlideShape,
  findSlideLayout,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeGradientFill,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';
import { renderEmfToSvg } from '../packages/preview/src/emf.ts';
import { buildPng } from './lib/build-png.ts';

const WORD_BYTES = 4;
const RECORD_HEADER_BYTES = 8;
const EMF_HEADER_BYTES = 88;
const EMF_SIGNATURE_OFFSET = 40;
const EMF_SIGNATURE = 0x464d4520;
const EMF_EXTENT = 100;
const record = (type: number, values: number[] = []): Uint8Array => {
  const bytes = new Uint8Array(RECORD_HEADER_BYTES + values.length * WORD_BYTES);
  const view = new DataView(bytes.buffer);
  [type, bytes.length, ...values].forEach((value, index) =>
    view.setUint32(index * WORD_BYTES, value, true),
  );
  return bytes;
};
const metafile = (records: Uint8Array[]): Uint8Array => {
  const header = new Uint8Array(EMF_HEADER_BYTES);
  const view = new DataView(header.buffer);
  view.setUint32(0, 1, true);
  view.setUint32(WORD_BYTES, EMF_HEADER_BYTES, true);
  view.setInt32(16, EMF_EXTENT, true);
  view.setInt32(20, EMF_EXTENT, true);
  view.setUint32(EMF_SIGNATURE_OFFSET, EMF_SIGNATURE, true);
  const parts = [header, ...records, record(14, [0, 16, 20])];
  const bytes = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    bytes.set(part, offset);
    offset += part.length;
  }
  new DataView(bytes.buffer).setUint32(48, bytes.length, true);
  new DataView(bytes.buffer).setUint32(52, parts.length, true);
  return bytes;
};
const pathRecords = [
  record(17, [8]),
  record(9, [EMF_EXTENT, EMF_EXTENT]),
  record(11, [EMF_EXTENT, EMF_EXTENT]),
  record(39, [1, 0, 0x00ffffff, 0]),
  record(37, [1]),
  record(19, [2]),
  record(59),
  record(27, [10, 10]),
  record(6, [0, 0, 90, 90, 2, 90, 10, 90, 90]),
  record(5, [0, 0, 90, 90, 3, 80, 90, 20, 90, 10, 10]),
  record(61),
  record(60),
  record(62, [0, 0, 90, 90]),
];

const blankSlide = async () => {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
  );
  const layout = findSlideLayout(pres, 'Blank')!;
  return { pres, slide: addSlide(pres, { layout }) };
};
const loadMetafile = async (bytes: Uint8Array) => {
  const { pres, slide } = await blankSlide();
  addSlideImage(slide, buildPng(1, 1, [0, 0, 0]), {
    x: inches(0),
    y: inches(0),
    w: inches(1),
    h: inches(1),
  });
  const { entries } = readZip(await savePresentation(pres));
  const modified = entries.map((entry) => {
    if (entry.name.startsWith('ppt/media/'))
      return { ...entry, name: entry.name.replace('.png', '.emf'), data: bytes };
    if (entry.name.endsWith('.rels') || entry.name === '[Content_Types].xml') {
      return {
        ...entry,
        data: new TextEncoder().encode(
          new TextDecoder()
            .decode(entry.data)
            .replaceAll('.png', '.emf')
            .replace(
              'Extension="png" ContentType="image/png"',
              'Extension="emf" ContentType="image/x-emf"',
            ),
        ),
      };
    }
    return entry;
  });
  const loaded = await loadPresentation(writeZip(modified));
  return { pres: loaded, slide: getSlides(loaded).at(-1)! };
};
const renderMetafile = async (bytes: Uint8Array) => {
  const { pres, slide } = await loadMetafile(bytes);
  return renderSlideToSvg(pres, slide);
};

describe('imported artwork', () => {
  it('位置が順不同のグラデーションを位置順に描画する', async () => {
    const { pres, slide } = await blankSlide();
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    setShapeGradientFill(shape, {
      angleDeg: 90,
      stops: [
        { offset: 0.52, color: '#003C7A' },
        { offset: 0, color: '#002060' },
        { offset: 1, color: '#01A8DC' },
      ],
    });
    const svg = renderSlideToSvg(pres, slide);
    const offsets = [...svg.matchAll(/<stop offset="([^"]+)"/g)].map((match) => Number(match[1]));
    expect(offsets).toEqual([0, 0.52, 1]);
  });

  it('EMF の塗りつぶしパスを透過 SVG 画像として描画する', async () => {
    const svg = await renderMetafile(metafile(pathRecords));
    expect(svg).not.toContain('data-pptx-fallback="image"');
    const encoded = svg.match(/href="data:image\/svg\+xml;base64,([^"]+)"/)?.[1];
    expect(encoded).toBeDefined();
    const artwork = Buffer.from(encoded!, 'base64').toString();
    expect(artwork).toContain('fill="#ffffff"');
    expect(artwork).toContain('L90 10');
    expect(artwork).toContain('C80 90 20 90 10 10');
  });

  it('Node の PNG 出力にも EMF が描画され、パスの外側は透明になる', async () => {
    const redPaths = pathRecords.map((bytes) =>
      new DataView(bytes.buffer).getUint32(0, true) === 39
        ? record(39, [1, 0, 0x000000ff, 0])
        : bytes,
    );
    const { pres, slide } = await loadMetafile(metafile(redPaths));
    const { image } = renderSlideToRgba(pres, slide, { width: 960 });
    const channels = 4;
    const pixel = (x: number, y: number) => [
      ...image.data.subarray(
        (y * image.width + x) * channels,
        (y * image.width + x + 1) * channels,
      ),
    ];
    expect(pixel(40, 40)).toEqual([255, 0, 0, 255]);
    expect(pixel(2, 2)).toEqual([255, 255, 255, 255]);
  });

  it('放射状グラデーションを中心から外側への位置順に並べる', async () => {
    const { pres, slide } = await blankSlide();
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    setShapeGradientFill(shape, {
      path: 'circle',
      stops: [
        { offset: 0.52, color: '#003C7A' },
        { offset: 0, color: '#002060' },
        { offset: 1, color: '#01A8DC' },
      ],
    });
    const svg = renderSlideToSvg(pres, slide);
    expect([...svg.matchAll(/<stop offset="([^"]+)"/g)].map((match) => Number(match[1]))).toEqual([
      0, 0.52, 1,
    ]);
  });

  it('未対応の描画命令では部分描画せず代替表示を維持する', async () => {
    expect(await renderMetafile(metafile([...pathRecords, record(84)]))).toContain(
      'data-pptx-fallback="image"',
    );
  });

  it('壊れたレコード長でも例外や無限ループにせず代替表示する', async () => {
    const bytes = metafile(pathRecords);
    new DataView(bytes.buffer).setUint32(EMF_HEADER_BYTES + WORD_BYTES, 0, true);
    expect(await renderMetafile(bytes)).toContain('data-pptx-fallback="image"');
  });
});

describe('EMF path decoding', () => {
  it('16 ビット座標の折れ線とベジェ曲線を符号付きで読み取る', () => {
    const packedPoint = (x: number, y: number) => ((y & 0xffff) << 16) | (x & 0xffff);
    const bytes = metafile([
      record(59),
      record(27, [0, 0]),
      record(89, [0, 0, 90, 90, 1, packedPoint(-10, 20)]),
      record(88, [
        0,
        0,
        90,
        90,
        3,
        packedPoint(-20, 30),
        packedPoint(40, -50),
        packedPoint(60, 70),
      ]),
      record(61),
      record(60),
      record(62, [0, 0, 90, 90]),
    ]);
    expect(renderEmfToSvg(bytes)).toContain('M0 0L-10 20C-20 30 40 -50 60 70Z');
  });

  it('ウィンドウとビューポートの原点・倍率を適用してクリップする', () => {
    const bytes = metafile([
      record(17, [8]),
      record(10, [10, 20]),
      record(12, [5, 6]),
      record(9, [100, 200]),
      record(11, [200, -400]),
      record(59),
      record(27, [10, 20]),
      record(6, [0, 0, 100, 100, 2, 30, 20, 30, 40]),
      record(61),
      record(60),
      record(67, [5]),
      ...pathRecords,
      record(75, [0, 5]),
      ...pathRecords,
    ]);
    const svg = renderEmfToSvg(bytes)!;
    expect(svg).toContain('d="M5 6L45 6L45 -34Z" clip-rule="evenodd"');
    expect(svg.match(/clip-path="url\(#emf-clip-0\)"/g)).toHaveLength(1);
    expect(svg.match(/fill="#ffffff"/g)).toHaveLength(2);
  });

  it.each([
    record(6, [0, 0, 90, 90, 0xffffffff]),
    record(5, [0, 0, 90, 90, 1, 10, 10]),
    record(27),
    record(70, [100]),
    record(9, [0, 100]),
  ])('不正なレコードのデータ長や座標を拒否する', (invalid) => {
    expect(
      renderEmfToSvg(
        metafile([
          record(59),
          record(27, [0, 0]),
          invalid,
          record(61),
          record(60),
          record(62, [0, 0, 90, 90]),
        ]),
      ),
    ).toBeNull();
  });

  it('Uint8Array のオフセットを尊重する', () => {
    const bytes = metafile(pathRecords);
    const padded = new Uint8Array(bytes.length + WORD_BYTES);
    padded.set(bytes, WORD_BYTES);
    expect(renderEmfToSvg(padded.subarray(WORD_BYTES))).toEqual(renderEmfToSvg(bytes));
  });

  it('末尾の欠落を拒否する', () => {
    const bytes = metafile(pathRecords);
    expect(renderEmfToSvg(bytes.subarray(0, bytes.length - WORD_BYTES))).toBeNull();
  });

  it('EMF+ コメント内の描画を無視して部分描画しない', () => {
    expect(
      renderEmfToSvg(metafile([...pathRecords, record(70, [WORD_BYTES, 0x2b464d45])])),
    ).toBeNull();
  });
});
