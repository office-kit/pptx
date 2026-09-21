import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import * as api from '../src/api/index.ts';
import { Presentation, Slide, Shape, compile } from '../packages/dsl/src/index.ts';

const box = { x: api.inches(1), y: api.inches(1), w: api.inches(3), h: api.inches(2) };
const fixture = () => {
  const pres = api.createPresentation();
  const slide = api.addBlankSlide(pres);
  const shape = api.addSlideTextBox(slide, { ...box, text: 'X' });
  const table = api.addSlideTable(slide, { ...box, rows: [['X']] });
  return { pres, slide, shape, cell: api.getTableCell(table, 0, 0) };
};

describe('authoring enum boundaries', () => {
  it.each([
    [
      'setShapeAlignment',
      ({ shape }: ReturnType<typeof fixture>) => {
        // @ts-expect-error JavaScript callers can supply unknown tokens.
        api.setShapeAlignment(shape, 'middle');
      },
    ],
    [
      'setTableCellAlignment',
      ({ cell }: ReturnType<typeof fixture>) => {
        // @ts-expect-error JavaScript callers can supply unknown tokens.
        api.setTableCellAlignment(cell, 'middle');
      },
    ],
    [
      'addSlideShape',
      ({ slide }: ReturnType<typeof fixture>) => {
        api.addSlideShape(slide, { ...box, preset: 'rounded-rect' });
      },
    ],
    [
      'setShapeStrokeDash',
      ({ shape }: ReturnType<typeof fixture>) => {
        // @ts-expect-error JavaScript callers can supply unknown tokens.
        api.setShapeStrokeDash(shape, 'dotted');
      },
    ],
    [
      'setShapeStrokeArrow',
      ({ shape }: ReturnType<typeof fixture>) => {
        // @ts-expect-error JavaScript callers can supply unknown tokens.
        api.setShapeStrokeArrow(shape, 'head', { type: 'arrowhead' });
      },
    ],
    [
      'setShapeTextAnchor',
      ({ shape }: ReturnType<typeof fixture>) => {
        // @ts-expect-error JavaScript callers can supply unknown tokens.
        api.setShapeTextAnchor(shape, 'ctr');
      },
    ],
    [
      'setShapeTextAutoFit',
      ({ shape }: ReturnType<typeof fixture>) => {
        // @ts-expect-error JavaScript callers can supply unknown tokens.
        api.setShapeTextAutoFit(shape, 'shrink');
      },
    ],
  ] as const)('%s rejects the issue reproduction without changing the slide', (name, apply) => {
    const f = fixture();
    const before = api.getSlideXmlString(f.slide);
    expect(() => apply(f)).toThrow(RangeError);
    expect(() => apply(f)).toThrow(new RegExp(`${name}: .*is not one of:`));
    expect(api.getSlideXmlString(f.slide)).toBe(before);
  });

  it('keeps known values and full-schema presets serializable', async () => {
    const { pres, slide, shape, cell } = fixture();
    api.setShapeAlignment(shape, 'center');
    api.setTableCellAlignment(cell, 'thaiDist');
    api.addSlideShape(slide, { ...box, preset: 'flowChartOfflineStorage' });
    api.setShapeStrokeDash(shape, 'sysDashDotDot');
    api.setShapeStrokeArrow(shape, 'head', { type: 'triangle', width: 'med', length: 'lg' });
    api.setShapeTextAnchor(shape, 'bottom');
    api.setShapeTextAutoFit(shape, 'normal');
    const reloaded = await api.loadPresentation(await api.savePresentation(pres));
    const xml = api.getSlideXmlString(api.getSlides(reloaded)[0]!);
    for (const token of [
      'algn="ctr"',
      'algn="thaiDist"',
      'prst="flowChartOfflineStorage"',
      'val="sysDashDotDot"',
      'type="triangle"',
      'w="med"',
      'len="lg"',
      'anchor="b"',
      '<a:normAutofit',
    ]) {
      expect(xml).toContain(token);
    }
  });

  it('rejects invalid presets through the DSL too', async () => {
    const deck = Presentation({
      children: Slide({
        children: Shape({ x: 1, y: 1, width: 2, height: 2, preset: 'rounded-rect' }),
      }),
    });
    await expect(compile(deck)).rejects.toThrow(
      /addSlideShape: preset: "rounded-rect" is not one of:/,
    );
  });
});

it('setShapeTextWrap validates wrap', () => {
  const { shape, slide } = fixture();
  expect(() => {
    // @ts-expect-error Exercise the JavaScript boundary.
    api.setShapeTextWrap(shape, 'bogus');
  }).toThrow(/setShapeTextWrap: .*is not one of:/);
  expect(() => api.setShapeTextWrap(shape, 'none')).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('wrap="none"');
});

it('setShapeTextDirection validates direction', () => {
  const { shape, slide } = fixture();
  expect(() => {
    // @ts-expect-error Exercise the JavaScript boundary.
    api.setShapeTextDirection(shape, 'bogus');
  }).toThrow(/setShapeTextDirection: .*is not one of:/);
  expect(() => api.setShapeTextDirection(shape, 'vert270')).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('vert="vert270"');
});

it('setTableCellTextDirection validates direction', () => {
  const { cell, slide } = fixture();
  expect(() => {
    // @ts-expect-error Exercise the JavaScript boundary.
    api.setTableCellTextDirection(cell, 'bogus');
  }).toThrow(/setTableCellTextDirection: .*is not one of:/);
  expect(() => api.setTableCellTextDirection(cell, 'vert270')).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('vert="vert270"');
});

it('setTableCellAnchor validates anchor', () => {
  const { cell, slide } = fixture();
  expect(() => {
    // @ts-expect-error Exercise the JavaScript boundary.
    api.setTableCellAnchor(cell, 'bogus');
  }).toThrow(/setTableCellAnchor: .*is not one of:/);
  expect(() => api.setTableCellAnchor(cell, 'center')).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('anchor="ctr"');
});

it('setShapeStrokeCap validates cap', () => {
  const { shape, slide } = fixture();
  expect(() => {
    // @ts-expect-error Exercise the JavaScript boundary.
    api.setShapeStrokeCap(shape, 'bogus');
  }).toThrow(/setShapeStrokeCap: .*is not one of:/);
  expect(() => api.setShapeStrokeCap(shape, 'rnd')).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('cap="rnd"');
});

it('setShapeStrokeJoin validates join', () => {
  const { shape, slide } = fixture();
  expect(() => {
    // @ts-expect-error Exercise the JavaScript boundary.
    api.setShapeStrokeJoin(shape, 'bogus');
  }).toThrow(/setShapeStrokeJoin: .*is not one of:/);
  expect(() => api.setShapeStrokeJoin(shape, 'bevel')).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('<a:bevel');
});

it('setShapeStrokeCompound validates compound', () => {
  const { shape, slide } = fixture();
  expect(() => {
    // @ts-expect-error Exercise the JavaScript boundary.
    api.setShapeStrokeCompound(shape, 'bogus');
  }).toThrow(/setShapeStrokeCompound: .*is not one of:/);
  expect(() => api.setShapeStrokeCompound(shape, 'dbl')).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('cmpd="dbl"');
});

it('setShapeStrokeArrow validates width', () => {
  const { shape, slide } = fixture();
  expect(() => {
    // @ts-expect-error Exercise the JavaScript boundary.
    api.setShapeStrokeArrow(shape, 'head', { type: 'triangle', width: 'bogus' });
  }).toThrow(/setShapeStrokeArrow: .*is not one of:/);
  expect(() =>
    api.setShapeStrokeArrow(shape, 'tail', { type: 'arrow', width: 'sm', length: 'lg' }),
  ).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('type="arrow"');
});

it('setShapeStrokeArrow validates length', () => {
  const { shape, slide } = fixture();
  expect(() => {
    // @ts-expect-error Exercise the JavaScript boundary.
    api.setShapeStrokeArrow(shape, 'head', { type: 'triangle', length: 'bogus' });
  }).toThrow(/setShapeStrokeArrow: .*is not one of:/);
  expect(() => api.setShapeStrokeArrow(shape, 'head', { type: 'none' })).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('type="none"');
});

it('setShapeBullets validates autoNum', () => {
  const { shape, slide } = fixture();
  expect(() => {
    api.setShapeBullets(shape, { autoNum: 'bogus' });
  }).toThrow(/setShapeBullets: .*is not one of:/);
  expect(() => api.setShapeBullets(shape, { autoNum: 'thaiNumPeriod' })).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('type="thaiNumPeriod"');
});

it('setShapeText validates bullets', () => {
  const { shape, slide } = fixture();
  api.setShapeText(shape, 'first\nsecond', { bullets: 'bullet' });
  const before = api.getSlideXmlString(slide);
  expect(() => {
    api.setShapeText(shape, 'new', { bullets: { autoNum: 'bogus' } });
  }).toThrow(/setShapeText: .*is not one of:/);
  expect(api.getSlideXmlString(slide)).toBe(before);
  expect(() => api.setShapeText(shape, 'new', { bullets: 'number' })).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('type="arabicPeriod"');
});

it('setTableCellBorders validates dash', () => {
  const { cell, slide } = fixture();
  expect(() => {
    api.setTableCellBorders(cell, { left: { dash: 'bogus' } });
  }).toThrow(/setTableCellBorders: .*is not one of:/);
  expect(() => api.setTableCellBorders(cell, { left: { dash: 'dashDot' } })).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('val="dashDot"');
});

it('setShapeTextFormat validates underline', () => {
  const { shape, slide } = fixture();
  expect(() => {
    api.setShapeTextFormat(shape, { underline: 'bogus' });
  }).toThrow(/setShapeTextFormat: .*is not one of:/);
  expect(() => api.setShapeTextFormat(shape, { underline: 'wavy' })).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('u="wavy"');
});

it('setTableCellTextFormat validates strike', () => {
  const { cell, slide } = fixture();
  expect(() => {
    api.setTableCellTextFormat(cell, { strike: 'bogus' });
  }).toThrow(/setTableCellTextFormat: .*is not one of:/);
  expect(() => api.setTableCellTextFormat(cell, { strike: 'dblStrike' })).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('strike="dblStrike"');
});

it('setShapeTextFormat validates cap', () => {
  const { shape, slide } = fixture();
  expect(() => {
    // @ts-expect-error Exercise the JavaScript boundary.
    api.setShapeTextFormat(shape, { cap: 'bogus' });
  }).toThrow(/setShapeTextFormat: .*is not one of:/);
  expect(() => api.setShapeTextFormat(shape, { cap: 'all' })).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('cap="all"');
});

it('setShapeParagraphs validates align', () => {
  const { shape, slide } = fixture();
  expect(() => {
    // @ts-expect-error Exercise the JavaScript boundary.
    api.setShapeParagraphs(shape, [{ align: 'bogus', runs: [] }]);
  }).toThrow(/setShapeParagraphs: .*is not one of:/);
  expect(() => api.setShapeParagraphs(shape, [{ align: 'justLow', runs: [] }])).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('algn="justLow"');
});

it('setTableCellParagraphs validates endFormat', () => {
  const { cell, slide } = fixture();
  expect(() => {
    api.setTableCellParagraphs(cell, [{ runs: [], endFormat: { underline: 'bogus' } }]);
  }).toThrow(/setTableCellParagraphs: .*is not one of:/);
  expect(() =>
    api.setTableCellParagraphs(cell, [{ runs: [], endFormat: { underline: 'dbl' } }]),
  ).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('u="dbl"');
});

it('setShapeGradientFill validates path', () => {
  const { shape, slide } = fixture();
  expect(() => {
    api.setShapeGradientFill(shape, {
      // @ts-expect-error Exercise the JavaScript boundary.
      path: 'bogus',
      stops: [
        { offset: 0, color: 'FF0000' },
        { offset: 1, color: 'FFFFFF' },
      ],
    });
  }).toThrow(/setShapeGradientFill: .*is not one of:/);
  expect(() =>
    api.setShapeGradientFill(shape, {
      path: 'circle',
      stops: [
        { offset: 0, color: 'FF0000' },
        { offset: 1, color: 'FFFFFF' },
      ],
    }),
  ).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('path="circle"');
});

it('setSlideTransition validates orientation', () => {
  const { slide } = fixture();
  expect(() => {
    // @ts-expect-error Exercise the JavaScript boundary.
    api.setSlideTransition(slide, { effect: 'split', orientation: 'bogus' });
  }).toThrow(/setSlideTransition: .*is not one of:/);
  expect(() =>
    api.setSlideTransition(slide, { effect: 'split', orientation: 'vert' }),
  ).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('orient="vert"');
});

it('setSlideTransition validates speed', () => {
  const { slide } = fixture();
  api.setSlideTransition(slide, { effect: 'split', orientation: 'vert', speed: 'slow' });
  const before = api.getSlideTransition(slide);
  expect(() => {
    // @ts-expect-error Exercise the JavaScript boundary.
    api.setSlideTransition(slide, { effect: 'fade', speed: 'bogus' });
  }).toThrow(/setSlideTransition: .*is not one of:/);
  expect(api.getSlideTransition(slide)).toEqual(before);
  expect(() => api.setSlideTransition(slide, { effect: 'fade', speed: 'fast' })).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('spd="fast"');
});

it('setShapeAnimation validates effect', () => {
  const { shape, slide } = fixture();
  expect(() => {
    // @ts-expect-error Exercise the JavaScript boundary.
    api.setShapeAnimation(shape, { effect: 'bogus' });
  }).toThrow(/setShapeAnimation: .*is not one of:/);
  expect(() => api.setShapeAnimation(shape, { effect: 'appear' })).not.toThrow();
  expect(api.getSlideXmlString(slide)).toContain('presetID="1"');
});

it('setSlideSize validates type', () => {
  const { pres } = fixture();
  expect(() => {
    api.setSlideSize(pres, { width: api.inches(10), height: api.inches(7.5), type: 'bogus' });
  }).toThrow(/setSlideSize: .*is not one of:/);
  expect(() => api.setSlideSize(pres, api.SLIDE_SIZE_4_3)).not.toThrow();
});

it('validates single-paragraph alignment and single-run formatting', () => {
  const { shape, slide } = fixture();
  // @ts-expect-error Exercise the JavaScript boundary.
  expect(() => api.setParagraphAlignment(shape, 0, 'bogus')).toThrow(
    /setParagraphAlignment: align: .*is not one of:/,
  );
  api.setParagraphAlignment(shape, 0, 'right');
  expect(api.getSlideXmlString(slide)).toContain('algn="r"');
  expect(() => api.setShapeRunFormat(shape, 0, 0, { underline: 'bogus' })).toThrow(
    /setShapeRunFormat: underline: .*is not one of:/,
  );
  api.setShapeRunFormat(shape, 0, 0, { underline: 'dbl' });
  expect(api.getSlideXmlString(slide)).toContain('u="dbl"');
});

it('validates paragraph auto-numbering before changing the paragraph', () => {
  const { shape, slide } = fixture();
  const before = api.getSlideXmlString(slide);
  expect(() => api.setParagraphBullet(shape, 0, { autoNum: 'bogus' })).toThrow(
    /setParagraphBullet: bullets.autoNum: .*is not one of:/,
  );
  expect(api.getSlideXmlString(slide)).toBe(before);
  api.setParagraphBullet(shape, 0, { autoNum: 'romanLcParenR' });
  expect(api.getSlideXmlString(slide)).toContain('type="romanLcParenR"');
});

it.skipIf(!isSchemaValidationAvailable())(
  'accepts every ST_ShapeType preset from the schema',
  () => {
    const schema = readFileSync(
      new URL(
        '../references/ecma-376-5th/ECMA-376/OfficeOpenXML-XMLSchema-Transitional/dml-main.xsd',
        import.meta.url,
      ),
      'utf8',
    );
    const shapeType = schema.match(
      /<xsd:simpleType name="ST_ShapeType">([\s\S]*?)<\/xsd:simpleType>/,
    )![1]!;
    const presets = [...shapeType.matchAll(/<xsd:enumeration value="([^"]+)"/g)].map(
      (match) => match[1]!,
    );
    expect(presets.length).toBeGreaterThan(180);
    const pres = api.createPresentation();
    const slide = api.addBlankSlide(pres);
    for (const preset of presets) api.addSlideShape(slide, { ...box, preset });
    expectSchemaValid(api.getSlideXmlString(slide), 'pml');
  },
);

it('validates shape-builder anchor and arrow end without changing existing shapes', () => {
  const { shape, slide } = fixture();
  const before = api.getSlideXmlString(slide);
  // @ts-expect-error Exercise the JavaScript boundary.
  expect(() => api.addSlideShape(slide, { ...box, preset: 'rect', textAnchor: 'bogus' })).toThrow(
    /addSlideShape: textAnchor: .*is not one of:/,
  );
  // @ts-expect-error Exercise the JavaScript boundary.
  expect(() => api.setShapeStrokeArrow(shape, 'bogus', { type: 'none' })).toThrow(
    /setShapeStrokeArrow: end: .*is not one of:/,
  );
  // @ts-expect-error Exercise the JavaScript boundary.
  expect(() => api.setShapeBullets(shape, 'bogus')).toThrow(
    /setShapeBullets: bullets: .*is not one of:/,
  );
  expect(api.getSlideXmlString(slide)).toBe(before);
  api.addSlideShape(slide, { ...box, preset: 'rect', text: 'X', textAnchor: 'ctr' });
  expect(api.getSlideXmlString(slide)).toContain('anchor="ctr"');
});

it('preserves JavaScript null clearing for underline, strike, and cap', () => {
  const { shape, slide } = fixture();
  api.setShapeTextFormat(shape, { underline: 'dbl', strike: 'dblStrike', cap: 'all' });
  const before = api.getSlideXmlString(slide);
  for (const attribute of ['u="dbl"', 'strike="dblStrike"', 'cap="all"'])
    expect(before).toContain(attribute);
  api.setShapeTextFormat(shape, {
    // @ts-expect-error Preserve the existing JavaScript null-clearing behavior.
    underline: null,
    // @ts-expect-error Preserve the existing JavaScript null-clearing behavior.
    strike: null,
    // @ts-expect-error Preserve the existing JavaScript null-clearing behavior.
    cap: null,
  });
  const after = api.getSlideXmlString(slide);
  for (const attribute of [' u=', ' strike=', ' cap=']) expect(after).not.toContain(attribute);
});

it.each(['alignment', 'textFormat'] as const)(
  'rejects invalid cell %s before creating a text body',
  (setter) => {
    const pres = api.createPresentation();
    const slide = api.addBlankSlide(pres);
    const table = api.addSlideTable(slide, { ...box, rows: [['a', 'b']] });
    api.mergeTableCells(table, { row: 0, col: 0, rowSpan: 1, colSpan: 2 }, { coveredText: 'drop' });
    const cell = api.getTableCell(table, 0, 1);
    api.setTableCellFill(cell, 'FF0000');
    const before = api.getSlideXmlString(slide);
    expect(() => {
      if (setter === 'alignment') {
        // @ts-expect-error Exercise the JavaScript boundary.
        api.setTableCellAlignment(cell, 'middle');
      } else {
        api.setTableCellTextFormat(cell, { underline: 'single' });
      }
    }).toThrow(RangeError);
    // A later successful setter must not commit a rejected setter's mutations.
    api.setTableCellFill(cell, 'FF0000');
    expect(api.getSlideXmlString(slide)).toBe(before);
  },
);
