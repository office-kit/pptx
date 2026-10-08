// The reference desktop app's default master structures, as XML strings: the default slide
// master and its eleven layouts (Insert Slide Master), the custom layout
// (Insert Layout), the placeholders the Master Layout and Insert Placeholder
// commands add, and the default notes and handout masters.
//
// The geometry is what the reference desktop app (2013 and later) uses for a 16:9 slide (12192000 × 6858000
// EMU) and a portrait 7.5 × 10 in notes page (6858000 × 9144000 EMU). For any
// other size the boxes are scaled per axis, which is how the reference desktop app keeps a
// master's layout when the slide or page size changes.
//
// Everything here is static boilerplate, so it is held as strings rather than
// built element by element; the results are validated against the ECMA-376
// XSDs in `test/fn-slide-masters.test.ts` and `test/fn-page-masters.test.ts`.

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const NS_DECLS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const SLIDE_W = 12192000;
const SLIDE_H = 6858000;
const PAGE_W = 6858000;
const PAGE_H = 9144000;

// Field ids only have to be GUIDs; the reference desktop app reuses one per field kind too.
const DATE_FIELD_ID = '{0B8E6F38-1A47-4C55-9C2B-3D5A1F0E7A01}';
const SLIDE_NUMBER_FIELD_ID = '{0B8E6F38-1A47-4C55-9C2B-3D5A1F0E7A02}';

export const STANDARD_CLR_MAP =
  'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"';

const escapeText = (value: string): string =>
  value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const escapeAttr = (value: string): string => escapeText(value).replace(/"/g, '&quot;');

/** A box in EMU. */
export interface Box {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

/** Scales a box authored on a `fromW × fromH` canvas to `toW × toH`. */
const scaleBox = (box: Box, from: { w: number; h: number }, to: { w: number; h: number }): Box => ({
  x: Math.round((box.x * to.w) / from.w),
  y: Math.round((box.y * to.h) / from.h),
  w: Math.round((box.w * to.w) / from.w),
  h: Math.round((box.h * to.h) / from.h),
});

const xfrm = (box: Box | null): string =>
  box === null
    ? ''
    : `<a:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.w}" cy="${box.h}"/></a:xfrm>`;

const run = (text: string): string =>
  `<a:r><a:rPr lang="en-US"/><a:t>${escapeText(text)}</a:t></a:r>`;

const LEVEL_PROMPTS = [
  'Click to edit Master text styles',
  'Second level',
  'Third level',
  'Fourth level',
  'Fifth level',
];

/** The five-level body prompt every master and layout text placeholder carries. */
const levelParagraphs = (): string =>
  LEVEL_PROMPTS.map((text, level) => `<a:p><a:pPr lvl="${level}"/>${run(text)}</a:p>`).join('');

const dateParagraph = (): string =>
  `<a:p><a:fld id="${DATE_FIELD_ID}" type="datetimeFigureOut"><a:rPr lang="en-US"/></a:fld><a:endParaRPr lang="en-US"/></a:p>`;
const slideNumberParagraph = (): string =>
  `<a:p><a:fld id="${SLIDE_NUMBER_FIELD_ID}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p>`;
const emptyParagraph = (): string => '<a:p><a:endParaRPr lang="en-US"/></a:p>';
const textParagraph = (text: string): string =>
  `<a:p>${run(text)}<a:endParaRPr lang="en-US"/></a:p>`;

interface PlaceholderSpec {
  readonly name: string;
  /** The `<p:ph>` attributes, already serialized. */
  readonly ph: string;
  readonly box: Box | null;
  /** Inner `<a:bodyPr>` attributes (and children, for autofit). */
  readonly bodyPr?: string | undefined;
  readonly lstStyle?: string | undefined;
  readonly paragraphs: string;
  /** Masters spell out their geometry and insets; layouts inherit them. */
  readonly full?: boolean;
  readonly locks?: string;
  readonly spPrExtra?: string;
}

const INSETS = 'vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"';

const placeholderXml = (id: number, spec: PlaceholderSpec): string => {
  const geometry = spec.full ? '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' : '';
  const spPrInner = `${xfrm(spec.box)}${geometry}${spec.spPrExtra ?? ''}`;
  const spPr = spPrInner === '' ? '<p:spPr/>' : `<p:spPr>${spPrInner}</p:spPr>`;
  const bodyPr = spec.bodyPr ?? (spec.full ? `<a:bodyPr ${INSETS}/>` : '<a:bodyPr/>');
  const lstStyle = spec.lstStyle ? `<a:lstStyle>${spec.lstStyle}</a:lstStyle>` : '<a:lstStyle/>';
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeAttr(`${spec.name} ${id - 1}`)}"/><p:cNvSpPr><a:spLocks ${spec.locks ?? 'noGrp="1"'}/></p:cNvSpPr><p:nvPr><p:ph ${spec.ph}/></p:nvPr></p:nvSpPr>${spPr}<p:txBody>${bodyPr}${lstStyle}${spec.paragraphs}</p:txBody></p:sp>`;
};

const SP_TREE_HEAD =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

const spTree = (specs: ReadonlyArray<PlaceholderSpec>): string =>
  `<p:spTree>${SP_TREE_HEAD}${specs.map((spec, i) => placeholderXml(i + 2, spec)).join('')}</p:spTree>`;

const solidTx1 = (tint?: number): string =>
  `<a:solidFill><a:schemeClr val="tx1">${tint === undefined ? '' : `<a:tint val="${tint}"/>`}</a:schemeClr></a:solidFill>`;

/** `lvl1pPr` … `lvl9pPr` built from one function of the level (0-based). */
const levels = (count: number, of: (level: number) => string): string =>
  Array.from({ length: count }, (_, level) => of(level)).join('');

// ---------------------------------------------------------------------------
// Slide master.

export type MasterPlaceholderType = 'title' | 'body' | 'dt' | 'ftr' | 'sldNum';

const MASTER_BOXES: Record<MasterPlaceholderType, Box> = {
  title: { x: 838200, y: 365125, w: 10515600, h: 1325563 },
  body: { x: 838200, y: 1825625, w: 10515600, h: 4351338 },
  dt: { x: 838200, y: 6356350, w: 2743200, h: 365125 },
  ftr: { x: 4038600, y: 6356350, w: 4114800, h: 365125 },
  sldNum: { x: 8610600, y: 6356350, w: 2743200, h: 365125 },
};

const footerStyle = (algn: 'l' | 'ctr' | 'r'): string =>
  `<a:lvl1pPr algn="${algn}"><a:defRPr sz="1200">${solidTx1(75000)}</a:defRPr></a:lvl1pPr>`;

const MASTER_SPECS: Record<MasterPlaceholderType, (box: Box) => PlaceholderSpec> = {
  title: (box) => ({
    name: 'Title Placeholder',
    ph: 'type="title"',
    box,
    full: true,
    bodyPr: `<a:bodyPr ${INSETS} anchor="ctr"><a:normAutofit/></a:bodyPr>`,
    paragraphs: textParagraph('Click to edit Master title style'),
  }),
  body: (box) => ({
    name: 'Text Placeholder',
    ph: 'type="body" idx="1"',
    box,
    full: true,
    bodyPr: `<a:bodyPr ${INSETS}><a:normAutofit/></a:bodyPr>`,
    paragraphs: levelParagraphs(),
  }),
  dt: (box) => ({
    name: 'Date Placeholder',
    ph: 'type="dt" sz="half" idx="2"',
    box,
    full: true,
    bodyPr: `<a:bodyPr ${INSETS} anchor="ctr"/>`,
    lstStyle: footerStyle('l'),
    paragraphs: dateParagraph(),
  }),
  ftr: (box) => ({
    name: 'Footer Placeholder',
    ph: 'type="ftr" sz="quarter" idx="3"',
    box,
    full: true,
    bodyPr: `<a:bodyPr ${INSETS} anchor="ctr"/>`,
    lstStyle: footerStyle('ctr'),
    paragraphs: emptyParagraph(),
  }),
  sldNum: (box) => ({
    name: 'Slide Number Placeholder',
    ph: 'type="sldNum" sz="quarter" idx="4"',
    box,
    full: true,
    bodyPr: `<a:bodyPr ${INSETS} anchor="ctr"/>`,
    lstStyle: footerStyle('r'),
    paragraphs: slideNumberParagraph(),
  }),
};

/** Master Layout's order, which is also the order the default master lists them. */
export const MASTER_PLACEHOLDER_TYPES: ReadonlyArray<MasterPlaceholderType> = [
  'title',
  'body',
  'dt',
  'ftr',
  'sldNum',
];

const slideScale = (size: { w: number; h: number }) => (box: Box) =>
  scaleBox(box, { w: SLIDE_W, h: SLIDE_H }, size);

/**
 * One default slide-master placeholder (`<p:sp>`), for Master Layout to put
 * back. `id` is its `cNvPr` id, unique within the master.
 */
export const masterPlaceholderXml = (
  type: MasterPlaceholderType,
  id: number,
  size: { w: number; h: number },
): string => {
  const xml = placeholderXml(id, MASTER_SPECS[type](slideScale(size)(MASTER_BOXES[type])));
  return xml.replace('<p:sp>', `<p:sp ${NS_DECLS}>`);
};

const MARGIN_STEP = 457200;
const BODY_SIZES = [2800, 2400, 2000, 1800, 1800, 1800, 1800, 1800, 1800];

const TX_STYLES =
  '<p:txStyles>' +
  `<p:titleStyle><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200">${solidTx1()}<a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle>` +
  `<p:bodyStyle>${levels(
    9,
    (level) =>
      `<a:lvl${level + 1}pPr marL="${228600 + level * MARGIN_STEP}" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="${level === 0 ? 1000 : 500}"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="${BODY_SIZES[level]}" kern="1200">${solidTx1()}<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${level + 1}pPr>`,
  )}</p:bodyStyle>` +
  `<p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${levels(
    9,
    (level) =>
      `<a:lvl${level + 1}pPr marL="${level * MARGIN_STEP}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200">${solidTx1()}<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${level + 1}pPr>`,
  )}</p:otherStyle>` +
  '</p:txStyles>';

/**
 * The default slide master Insert Slide Master adds. `layouts` are the
 * `<p:sldLayoutId>` entries (`id`, `rId`) in order. The reference desktop app marks a
 * master it inserts as preserved, so it survives having no slides.
 */
export const defaultSlideMasterXml = (
  size: { w: number; h: number },
  layouts: ReadonlyArray<{ id: number; rId: string }>,
): string => {
  const scale = slideScale(size);
  const specs = MASTER_PLACEHOLDER_TYPES.map((type) =>
    MASTER_SPECS[type](scale(MASTER_BOXES[type])),
  );
  const ids = layouts.map((l) => `<p:sldLayoutId id="${l.id}" r:id="${l.rId}"/>`).join('');
  return `${XML_DECL}<p:sldMaster ${NS_DECLS} preserve="1"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>${spTree(specs)}</p:cSld><p:clrMap ${STANDARD_CLR_MAP}/><p:sldLayoutIdLst>${ids}</p:sldLayoutIdLst>${TX_STYLES}</p:sldMaster>`;
};

// ---------------------------------------------------------------------------
// Slide layouts.

const LAYOUT_FOOTERS: ReadonlyArray<PlaceholderSpec> = [
  {
    name: 'Date Placeholder',
    ph: 'type="dt" sz="half" idx="10"',
    box: null,
    paragraphs: dateParagraph(),
  },
  {
    name: 'Footer Placeholder',
    ph: 'type="ftr" sz="quarter" idx="11"',
    box: null,
    paragraphs: emptyParagraph(),
  },
  {
    name: 'Slide Number Placeholder',
    ph: 'type="sldNum" sz="quarter" idx="12"',
    box: null,
    paragraphs: slideNumberParagraph(),
  },
];

/** The footers a layout's Footers checkbox puts back (`<p:sp>` each, ids from `firstId`). */
export const layoutFooterPlaceholdersXml = (
  firstId: number,
  idx: readonly [number, number, number],
): string[] =>
  LAYOUT_FOOTERS.map((spec, i) =>
    placeholderXml(firstId + i, {
      ...spec,
      ph: spec.ph.replace(/idx="\d+"/, `idx="${idx[i]}"`),
    }).replace('<p:sp>', `<p:sp ${NS_DECLS}>`),
  );

const LAYOUT_TITLE: PlaceholderSpec = {
  name: 'Title',
  ph: 'type="title"',
  box: null,
  paragraphs: textParagraph('Click to edit Master title style'),
};

/** The title a layout's Title checkbox puts back. */
export const layoutTitlePlaceholderXml = (id: number): string =>
  placeholderXml(id, LAYOUT_TITLE).replace('<p:sp>', `<p:sp ${NS_DECLS}>`);

const centeredLevels = (sizes: ReadonlyArray<number>): string =>
  levels(
    9,
    (level) =>
      `<a:lvl${level + 1}pPr marL="${level * MARGIN_STEP}" indent="0" algn="ctr"><a:buNone/><a:defRPr sz="${sizes[Math.min(level, sizes.length - 1)]}"/></a:lvl${level + 1}pPr>`,
  );

const plainLevels = (
  sizes: ReadonlyArray<number>,
  options: { bold?: boolean; tint?: number; bullets?: boolean } = {},
): string =>
  levels(9, (level) => {
    const size = sizes[Math.min(level, sizes.length - 1)];
    const indent = options.bullets ? '' : ` marL="${level * MARGIN_STEP}" indent="0"`;
    const fill = options.tint === undefined ? '' : solidTx1(options.tint);
    const bold = options.bold ? ' b="1"' : '';
    return `<a:lvl${level + 1}pPr${indent}>${options.bullets ? '' : '<a:buNone/>'}<a:defRPr sz="${size}"${bold}${fill === '' ? '/>' : `>${fill}</a:defRPr>`}</a:lvl${level + 1}pPr>`;
  });

const titleAt = (
  box: Box,
  size?: number,
  anchorBottom = false,
  type = 'title',
): PlaceholderSpec => ({
  name: 'Title',
  ph: `type="${type}"`,
  box,
  bodyPr: anchorBottom ? '<a:bodyPr anchor="b"/>' : undefined,
  lstStyle:
    size === undefined
      ? undefined
      : `<a:lvl1pPr${type === 'ctrTitle' ? ' algn="ctr"' : ''}><a:defRPr sz="${size}"/></a:lvl1pPr>`,
  paragraphs: textParagraph('Click to edit Master title style'),
});

const content = (
  name: string,
  ph: string,
  box: Box | null,
  lstStyle?: string,
  bodyPr?: string,
): PlaceholderSpec => ({
  name,
  ph,
  box,
  bodyPr,
  lstStyle,
  paragraphs: levelParagraphs(),
});

const CAPTION_TITLE: Box = { x: 839788, y: 457200, w: 3932237, h: 1600200 };
const CAPTION_BODY: Box = { x: 839788, y: 2057400, w: 3932237, h: 3811588 };
const CAPTION_RIGHT: Box = { x: 5183188, y: 987425, w: 6172200, h: 4873625 };
const CAPTION_TEXT = plainLevels([1600, 1400, 1200, 1000]);

interface LayoutSpec {
  readonly name: string;
  readonly type: string;
  readonly placeholders: ReadonlyArray<PlaceholderSpec>;
}

// The eleven layouts of the default master, in the reference desktop app's order.
const DEFAULT_LAYOUTS: ReadonlyArray<LayoutSpec> = [
  {
    name: 'Title Slide',
    type: 'title',
    placeholders: [
      titleAt({ x: 1524000, y: 1122363, w: 9144000, h: 2387600 }, 6000, true, 'ctrTitle'),
      {
        name: 'Subtitle',
        ph: 'type="subTitle" idx="1"',
        box: { x: 1524000, y: 3602038, w: 9144000, h: 1655762 },
        lstStyle: centeredLevels([2400, 2000, 1800, 1600]),
        paragraphs: textParagraph('Click to edit Master subtitle style'),
      },
    ],
  },
  {
    name: 'Title and Content',
    type: 'obj',
    placeholders: [LAYOUT_TITLE, content('Content Placeholder', 'idx="1"', null)],
  },
  {
    name: 'Section Header',
    type: 'secHead',
    placeholders: [
      titleAt({ x: 831850, y: 1709738, w: 10515600, h: 2852737 }, 6000, true),
      {
        name: 'Text Placeholder',
        ph: 'type="body" idx="1"',
        box: { x: 831850, y: 4589463, w: 10515600, h: 1500187 },
        lstStyle: plainLevels([2400, 2000, 1800, 1600], { tint: 82000 }),
        paragraphs: levelParagraphs(),
      },
    ],
  },
  {
    name: 'Two Content',
    type: 'twoObj',
    placeholders: [
      LAYOUT_TITLE,
      content('Content Placeholder', 'sz="half" idx="1"', {
        x: 838200,
        y: 1825625,
        w: 5181600,
        h: 4351338,
      }),
      content('Content Placeholder', 'sz="half" idx="2"', {
        x: 6172200,
        y: 1825625,
        w: 5181600,
        h: 4351338,
      }),
    ],
  },
  {
    name: 'Comparison',
    type: 'twoTxTwoObj',
    placeholders: [
      { ...LAYOUT_TITLE, box: { x: 839788, y: 365125, w: 10515600, h: 1325563 } },
      content(
        'Text Placeholder',
        'type="body" idx="1"',
        { x: 839788, y: 1681163, w: 5157787, h: 823912 },
        plainLevels([2400, 2000, 1800, 1600], { bold: true }),
        '<a:bodyPr anchor="b"/>',
      ),
      content('Content Placeholder', 'sz="half" idx="2"', {
        x: 839788,
        y: 2505075,
        w: 5157787,
        h: 3684588,
      }),
      content(
        'Text Placeholder',
        'type="body" sz="quarter" idx="3"',
        { x: 6172200, y: 1681163, w: 5183188, h: 823912 },
        plainLevels([2400, 2000, 1800, 1600], { bold: true }),
        '<a:bodyPr anchor="b"/>',
      ),
      content('Content Placeholder', 'sz="quarter" idx="4"', {
        x: 6172200,
        y: 2505075,
        w: 5183188,
        h: 3684588,
      }),
    ],
  },
  { name: 'Title Only', type: 'titleOnly', placeholders: [LAYOUT_TITLE] },
  { name: 'Blank', type: 'blank', placeholders: [] },
  {
    name: 'Content with Caption',
    type: 'objTx',
    placeholders: [
      titleAt(CAPTION_TITLE, 3200, true),
      content(
        'Content Placeholder',
        'idx="1"',
        CAPTION_RIGHT,
        plainLevels([3200, 2800, 2400, 2000], { bullets: true }),
      ),
      content('Text Placeholder', 'type="body" sz="half" idx="2"', CAPTION_BODY, CAPTION_TEXT),
    ],
  },
  {
    name: 'Picture with Caption',
    type: 'picTx',
    placeholders: [
      titleAt(CAPTION_TITLE, 3200, true),
      {
        name: 'Picture Placeholder',
        ph: 'type="pic" idx="1"',
        box: CAPTION_RIGHT,
        bodyPr: '<a:bodyPr anchor="t"/>',
        lstStyle: plainLevels([3200, 2800, 2400, 2000]),
        paragraphs: emptyParagraph(),
      },
      content('Text Placeholder', 'type="body" sz="half" idx="2"', CAPTION_BODY, CAPTION_TEXT),
    ],
  },
  {
    name: 'Title and Vertical Text',
    type: 'vertTx',
    placeholders: [
      LAYOUT_TITLE,
      content(
        'Vertical Text Placeholder',
        'type="body" orient="vert" idx="1"',
        null,
        undefined,
        '<a:bodyPr vert="eaVert"/>',
      ),
    ],
  },
  {
    name: 'Vertical Title and Text',
    type: 'vertTitleAndTx',
    placeholders: [
      {
        name: 'Vertical Title',
        ph: 'type="title" orient="vert"',
        box: { x: 8724900, y: 365125, w: 2628900, h: 5811838 },
        bodyPr: '<a:bodyPr vert="eaVert"/>',
        paragraphs: textParagraph('Click to edit Master title style'),
      },
      content(
        'Vertical Text Placeholder',
        'type="body" orient="vert" idx="1"',
        { x: 838200, y: 365125, w: 7734300, h: 5811838 },
        undefined,
        '<a:bodyPr vert="eaVert"/>',
      ),
    ],
  },
];

const layoutXml = (
  spec: LayoutSpec,
  size: { w: number; h: number },
  extraAttrs: string,
): string => {
  const scale = slideScale(size);
  const placeholders = [...spec.placeholders, ...LAYOUT_FOOTERS].map((ph) =>
    ph.box === null ? ph : { ...ph, box: scale(ph.box) },
  );
  return `${XML_DECL}<p:sldLayout ${NS_DECLS}${extraAttrs}><p:cSld name="${escapeAttr(spec.name)}">${spTree(placeholders)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
};

/** The default master's eleven layouts, in order, for a slide of `size`. */
export const defaultSlideLayoutsXml = (size: { w: number; h: number }): string[] =>
  DEFAULT_LAYOUTS.map((spec) => layoutXml(spec, size, ` type="${spec.type}" preserve="1"`));

/**
 * The layout Insert Layout adds: a title and the three footers, named
 * "Custom Layout" (callers number repeats as the reference desktop app does).
 */
export const customSlideLayoutXml = (name: string, size: { w: number; h: number }): string =>
  layoutXml(
    { name, type: 'cust', placeholders: [LAYOUT_TITLE] },
    size,
    ' preserve="1" userDrawn="1"',
  );

// ---------------------------------------------------------------------------
// Insert Placeholder.

/** Insert Placeholder's menu, in the reference desktop app's order. */
export type LayoutPlaceholderKind =
  | 'content'
  | 'verticalContent'
  | 'text'
  | 'verticalText'
  | 'picture'
  | 'chart'
  | 'table'
  | 'smartArt'
  | 'media'
  | 'onlineImage';

const INSERTED: Record<
  LayoutPlaceholderKind,
  { name: string; ph: string; prompt: string | null; vertical?: boolean }
> = {
  content: { name: 'Content Placeholder', ph: '', prompt: null },
  verticalContent: {
    name: 'Vertical Content Placeholder',
    ph: 'orient="vert" ',
    prompt: null,
    vertical: true,
  },
  text: { name: 'Text Placeholder', ph: 'type="body" ', prompt: null },
  verticalText: {
    name: 'Vertical Text Placeholder',
    ph: 'type="body" orient="vert" ',
    prompt: null,
    vertical: true,
  },
  picture: { name: 'Picture Placeholder', ph: 'type="pic" ', prompt: 'Picture' },
  chart: { name: 'Chart Placeholder', ph: 'type="chart" ', prompt: 'Chart' },
  table: { name: 'Table Placeholder', ph: 'type="tbl" ', prompt: 'Table' },
  smartArt: { name: 'SmartArt Placeholder', ph: 'type="dgm" ', prompt: 'SmartArt' },
  media: { name: 'Media Placeholder', ph: 'type="media" ', prompt: 'Media' },
  // The reference desktop app keeps the pre-2013 `clipArt` token for its Online Image placeholder.
  onlineImage: { name: 'Online Image Placeholder', ph: 'type="clipArt" ', prompt: 'Online Image' },
};

/** A placeholder Insert Placeholder draws on a layout. */
export const insertedPlaceholderXml = (
  kind: LayoutPlaceholderKind,
  id: number,
  idx: number,
  box: Box,
): string => {
  const spec = INSERTED[kind];
  return placeholderXml(id, {
    name: spec.name,
    ph: `${spec.ph}sz="quarter" idx="${idx}"`,
    box,
    bodyPr: spec.vertical ? '<a:bodyPr vert="eaVert"/>' : undefined,
    paragraphs: spec.prompt === null ? levelParagraphs() : textParagraph(spec.prompt),
  }).replace('<p:sp>', `<p:sp ${NS_DECLS}>`);
};

// ---------------------------------------------------------------------------
// Notes and handout masters.

export type NotesMasterPlaceholderType = 'hdr' | 'dt' | 'sldImg' | 'body' | 'ftr' | 'sldNum';
export type HandoutMasterPlaceholderType = 'hdr' | 'dt' | 'ftr' | 'sldNum';

const CORNERS: Record<HandoutMasterPlaceholderType, Box> = {
  hdr: { x: 0, y: 0, w: 2971800, h: 458788 },
  dt: { x: 3884613, y: 0, w: 2971800, h: 458788 },
  ftr: { x: 0, y: 8685213, w: 2971800, h: 458788 },
  sldNum: { x: 3884613, y: 8685213, w: 2971800, h: 458788 },
};
const NOTES_BODY: Box = { x: 685800, y: 4400550, w: 5486400, h: 3600450 };

const pageStyle = (algn: 'l' | 'r'): string =>
  `<a:lvl1pPr algn="${algn}"><a:defRPr sz="1200"/></a:lvl1pPr>`;

const cornerSpec = (
  type: HandoutMasterPlaceholderType,
  idx: number | null,
  box: Box,
): PlaceholderSpec => {
  const idxAttr = idx === null ? '' : ` idx="${idx}"`;
  switch (type) {
    case 'hdr':
      return {
        name: 'Header Placeholder',
        ph: `type="hdr" sz="quarter"${idxAttr}`,
        box,
        full: true,
        lstStyle: pageStyle('l'),
        paragraphs: emptyParagraph(),
      };
    case 'dt':
      return {
        name: 'Date Placeholder',
        ph: `type="dt" sz="quarter"${idxAttr}`,
        box,
        full: true,
        lstStyle: pageStyle('r'),
        paragraphs: dateParagraph(),
      };
    case 'ftr':
      return {
        name: 'Footer Placeholder',
        ph: `type="ftr" sz="quarter"${idxAttr}`,
        box,
        full: true,
        bodyPr: `<a:bodyPr ${INSETS} anchor="b"/>`,
        lstStyle: pageStyle('l'),
        paragraphs: emptyParagraph(),
      };
    case 'sldNum':
      return {
        name: 'Slide Number Placeholder',
        ph: `type="sldNum" sz="quarter"${idxAttr}`,
        box,
        full: true,
        bodyPr: `<a:bodyPr ${INSETS} anchor="b"/>`,
        lstStyle: pageStyle('r'),
        paragraphs: slideNumberParagraph(),
      };
  }
};

/**
 * Where the notes page's slide image goes. The reference desktop app keeps the slide's aspect
 * ratio: a widescreen slide fills the 6 in width 1.25 in from the top, a 4:3
 * one fills 3.75 in of height 0.75 in from the top, centred.
 */
const slideImageBox = (slide: { w: number; h: number }): Box => {
  const aspect = slide.w / slide.h;
  if (aspect >= 1.5) {
    const w = 5486400;
    return { x: 685800, y: 1143000, w, h: Math.round(w / aspect) };
  }
  const h = 3429000;
  const w = Math.round(h * aspect);
  return { x: Math.round((PAGE_W - w) / 2), y: 685800, w, h };
};

const NOTES_IDX: Record<NotesMasterPlaceholderType, number | null> = {
  hdr: null,
  dt: 1,
  sldImg: 2,
  body: 3,
  ftr: 4,
  sldNum: 5,
};
const HANDOUT_IDX: Record<HandoutMasterPlaceholderType, number | null> = {
  hdr: null,
  dt: 1,
  ftr: 2,
  sldNum: 3,
};

export const NOTES_MASTER_PLACEHOLDER_TYPES: ReadonlyArray<NotesMasterPlaceholderType> = [
  'hdr',
  'dt',
  'sldImg',
  'body',
  'ftr',
  'sldNum',
];
export const HANDOUT_MASTER_PLACEHOLDER_TYPES: ReadonlyArray<HandoutMasterPlaceholderType> = [
  'hdr',
  'dt',
  'ftr',
  'sldNum',
];

const pageScale = (page: { w: number; h: number }) => (box: Box) =>
  scaleBox(box, { w: PAGE_W, h: PAGE_H }, page);

const notesSpec = (
  type: NotesMasterPlaceholderType,
  page: { w: number; h: number },
  slide: { w: number; h: number },
): PlaceholderSpec => {
  const scale = pageScale(page);
  if (type === 'sldImg') {
    return {
      name: 'Slide Image Placeholder',
      ph: `type="sldImg" idx="${NOTES_IDX.sldImg}"`,
      box: scale(slideImageBox(slide)),
      full: true,
      locks: 'noGrp="1" noRot="1" noChangeAspect="1"',
      spPrExtra:
        '<a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln>',
      bodyPr: `<a:bodyPr ${INSETS} anchor="ctr"/>`,
      paragraphs: emptyParagraph(),
    };
  }
  if (type === 'body') {
    return {
      name: 'Notes Placeholder',
      ph: `type="body" sz="quarter" idx="${NOTES_IDX.body}"`,
      box: scale(NOTES_BODY),
      full: true,
      paragraphs: levelParagraphs(),
    };
  }
  return cornerSpec(type, NOTES_IDX[type], scale(CORNERS[type]));
};

/** One default notes-master placeholder, for the Notes Master checkboxes to put back. */
export const notesMasterPlaceholderXml = (
  type: NotesMasterPlaceholderType,
  id: number,
  page: { w: number; h: number },
  slide: { w: number; h: number },
): string =>
  placeholderXml(id, notesSpec(type, page, slide)).replace('<p:sp>', `<p:sp ${NS_DECLS}>`);

/** One default handout-master placeholder. */
export const handoutMasterPlaceholderXml = (
  type: HandoutMasterPlaceholderType,
  id: number,
  page: { w: number; h: number },
): string =>
  placeholderXml(id, cornerSpec(type, HANDOUT_IDX[type], pageScale(page)(CORNERS[type]))).replace(
    '<p:sp>',
    `<p:sp ${NS_DECLS}>`,
  );

const NOTES_STYLE = `<p:notesStyle>${levels(
  9,
  (level) =>
    `<a:lvl${level + 1}pPr marL="${level * MARGIN_STEP}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200">${solidTx1()}<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${level + 1}pPr>`,
)}</p:notesStyle>`;

const PAGE_BACKGROUND = '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>';

/** The reference desktop app's default notes master for a page of `page` and slides of `slide`. */
export const defaultNotesMasterXml = (
  page: { w: number; h: number },
  slide: { w: number; h: number },
): string => {
  const specs = NOTES_MASTER_PLACEHOLDER_TYPES.map((type) => notesSpec(type, page, slide));
  return `${XML_DECL}<p:notesMaster ${NS_DECLS}><p:cSld>${PAGE_BACKGROUND}${spTree(specs)}</p:cSld><p:clrMap ${STANDARD_CLR_MAP}/>${NOTES_STYLE}</p:notesMaster>`;
};

/** The reference desktop app's default handout master for a page of `page`. */
export const defaultHandoutMasterXml = (page: { w: number; h: number }): string => {
  const scale = pageScale(page);
  const specs = HANDOUT_MASTER_PLACEHOLDER_TYPES.map((type) =>
    cornerSpec(type, HANDOUT_IDX[type], scale(CORNERS[type])),
  );
  return `${XML_DECL}<p:handoutMaster ${NS_DECLS}><p:cSld>${PAGE_BACKGROUND}${spTree(specs)}</p:cSld><p:clrMap ${STANDARD_CLR_MAP}/></p:handoutMaster>`;
};
