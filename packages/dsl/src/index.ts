import * as api from '@office-kit/pptx';
import {
  created,
  textContent,
  literal,
  node,
  requireSlide,
  visit,
  type Children,
  type Context,
  type Node,
  type PresentationNode,
  type RawContext,
} from './model.ts';
export type { Child, Children, Node, PresentationNode, RawContext } from './model.ts';
export { Fragment } from './jsx-runtime.ts';

export type PresentationProps = Children &
  (
    | { source: api.PresentationInput; size?: never }
    | { source?: never; size?: api.PresentationSize }
  ) & {
    theme?: Parameters<typeof api.setPresentationTheme>[1];
    /** Compose replaces the slide sequence; edit retains the original sequence. */
    mode?: 'compose' | 'edit';
  };
export function Presentation(props: PresentationProps): PresentationNode {
  return {
    kind: 'Presentation',
    evaluate: () => {
      throw new Error('Presentation must be the root.');
    },
    async compile() {
      if (props.source !== undefined && props.size !== undefined)
        throw new Error('Presentation source retains its original size; do not specify size.');
      const presentation =
        props.source === undefined
          ? api.createPresentation(props.size === undefined ? {} : { size: props.size })
          : await api.loadPresentation(props.source);
      if (props.theme) api.setPresentationTheme(presentation, props.theme);
      const originals = [...api.getSlides(presentation)];
      const context: Context = {
        presentation,
        originals,
        mode: props.mode ?? (props.source === undefined ? 'compose' : 'edit'),
        scope: 'presentation',
        deferred: [],
      };
      await visit(props.children, context);
      for (const apply of context.deferred) await apply();
      if (context.mode === 'compose') {
        for (const slide of originals) api.removeSlide(presentation, slide);
      }
      return presentation;
    },
  };
}
export async function compile(root: Node): Promise<api.PresentationData> {
  if (!('compile' in root) || typeof root.compile !== 'function')
    throw new Error('Export a Presentation as the root.');
  return root.compile();
}

export type LayoutRef = { name: string } | { part: string } | { type: string };
export type SlideRef = { index: number } | { part: string } | { title: string };
export type Target =
  | { name: string }
  | { id: number }
  | { placeholder: { type?: string; idx?: number } };
export type SlideProps = Children &
  ({ from?: SlideRef; target?: never } | { from?: never; target: SlideRef }) & {
    layout?: LayoutRef;
    background?: api.Color;
    notes?: string;
  };
function selectSlide(context: Context, ref: SlideRef): api.SlideData {
  const matches = context.originals.filter((slide, index) =>
    'index' in ref
      ? index === ref.index
      : 'part' in ref
        ? api.getSlidePartName(slide) === ref.part
        : api.getSlideTitle(slide) === ref.title,
  );
  if (matches.length !== 1)
    throw new Error(`Slide reference matched ${matches.length} slides: ${JSON.stringify(ref)}`);
  return matches[0]!;
}
function selectLayout(presentation: api.PresentationData, ref: LayoutRef): api.SlideLayoutData {
  const matches = api
    .getSlideLayouts(presentation)
    .filter((layout) =>
      'name' in ref
        ? api.getSlideLayoutName(layout) === ref.name
        : 'part' in ref
          ? api.getSlideLayoutPartName(layout) === ref.part
          : api.getSlideLayoutType(layout) === ref.type,
    );
  if (matches.length !== 1)
    throw new Error(`Layout reference matched ${matches.length} layouts: ${JSON.stringify(ref)}`);
  return matches[0]!;
}
export function Slide(props: SlideProps): Node {
  return node('Slide', async (context) => {
    if (context.scope !== 'presentation') throw new Error('Slide must be inside Presentation.');
    if (props.target && context.mode !== 'edit')
      throw new Error('Slide target requires Presentation mode="edit".');
    if (props.from && props.target) throw new Error('Slide cannot specify both from and target.');
    const layout = props.layout ? selectLayout(context.presentation, props.layout) : undefined;
    const slide = props.target
      ? selectSlide(context, props.target)
      : props.from
        ? api.duplicateSlide(context.presentation, selectSlide(context, props.from))
        : layout
          ? api.addSlide(context.presentation, { layout })
          : api.addBlankSlide(context.presentation);
    if (layout && (props.from || props.target)) api.setSlideLayout(slide, layout);
    if (props.background !== undefined) api.setSlideBackground(slide, props.background);
    if (props.notes !== undefined) api.setSlideNotes(slide, props.notes);
    await visit(props.children, { ...context, slide, scope: 'slide' });
  });
}
// A format read back from the cell carries colors as plain strings, since a
// deck can hold a scheme token outside the theme. Re-applying it has to go
// through the same check the writer would run.
const writableFormat = (format: api.ReadTextFormat): api.TextFormat => {
  const { color, highlight, ...rest } = format;
  return {
    ...rest,
    ...(color == null ? {} : { color: api.asColor(color) }),
    ...(highlight == null ? {} : { highlight: api.asColor(highlight) }),
  };
};

export type RawProps =
  | { scope?: undefined; apply: (context: RawContext) => void | Promise<void> }
  | {
      scope: 'presentation';
      apply: (context: Pick<RawContext, 'presentation'>) => void | Promise<void>;
    }
  | {
      scope: 'slide';
      apply: (context: RawContext & { slide: api.SlideData }) => void | Promise<void>;
    }
  | { scope: 'shape'; apply: (context: Required<RawContext>) => void | Promise<void> };

export function Raw(props: RawProps): Node {
  return node('Raw', (context) => {
    if (props.scope !== undefined && props.scope !== context.scope) {
      throw new Error(`Raw scope="${props.scope}" cannot run inside ${context.scope}.`);
    }
    context.deferred.push(() => {
      if (props.scope === 'shape') {
        if (!context.shape || !context.slide) throw new Error('Raw requires a shape context.');
        return props.apply({
          presentation: context.presentation,
          slide: context.slide,
          shape: context.shape,
        });
      }
      if (props.scope === 'slide') {
        if (!context.slide) throw new Error('Raw requires a slide context.');
        return props.apply({ ...context, slide: context.slide });
      }
      return props.apply(context);
    });
  });
}
export interface Bounds {
  x: number;
  y: number;
  width: number;
  height: number;
}
export interface ShapeStyle {
  name?: string;
  fill?: api.Color | api.GradientFillOptions;
  stroke?: { color: api.Color; width?: number } | false;
  rotation?: number;
  shadow?: api.ShadowOptions;
  glow?: api.GlowOptions;
  click?: api.ShapeClickAction;
}
function bounds(props: Bounds) {
  return {
    x: api.inches(props.x),
    y: api.inches(props.y),
    w: api.inches(props.width),
    h: api.inches(props.height),
  };
}
function style(shape: api.SlideShapeData, props: ShapeStyle) {
  if (props.fill !== undefined) {
    if (typeof props.fill === 'string') api.setShapeFill(shape, props.fill);
    else api.setShapeGradientFill(shape, props.fill);
  }
  if (props.stroke === false) api.setShapeNoStroke(shape);
  else if (props.stroke)
    api.setShapeStroke(shape, {
      color: props.stroke.color,
      widthEmu: api.pt(props.stroke.width ?? 1),
    });
  if (props.rotation !== undefined) api.setShapeRotation(shape, props.rotation);
  if (props.shadow) api.setShapeShadow(shape, props.shadow);
  if (props.glow) api.setShapeGlow(shape, props.glow);
  if (props.click) api.setShapeClickAction(shape, props.click);
}
/** A paragraph's indent level: 0 is the top, 8 the deepest PowerPoint has. */
export type TextLevel = 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8;
export interface TextParagraph extends api.ParagraphSpec {
  /** This paragraph's own bullet; wins over `bullets`. */
  bullet?: api.BulletStyle;
  level?: TextLevel;
}
export interface TextProps extends Bounds, ShapeStyle, api.TextFormat, Children {
  paragraphs?: readonly TextParagraph[];
  align?: api.ParagraphAlignment;
  anchor?: api.TextAnchor;
  autoFit?: api.TextAutoFit;
  /** Bullet or numbering for every paragraph; a newline in the text starts a paragraph. */
  bullets?: api.BulletStyle;
  /** Space before / after every paragraph, in points. */
  paragraphSpacing?: { before?: number; after?: number };
}
export function Text(props: TextProps): Node {
  return node('Text', async (context) => {
    const slide = requireSlide(context);
    const content = textContent(props.children);
    if (props.paragraphs && content.text.trim())
      throw new Error('Text accepts either text children or paragraphs.');
    const shape = created(
      context,
      api.addSlideTextBox(slide, {
        ...bounds(props),
        ...(props.name ? { name: props.name } : {}),
        text: content.text,
      }),
    );
    if (props.paragraphs) api.setShapeParagraphs(shape, props.paragraphs);
    api.setShapeTextFormat(shape, props);
    if (props.align) api.setShapeAlignment(shape, props.align);
    if (props.anchor) api.setShapeTextAnchor(shape, props.anchor);
    if (props.autoFit) api.setShapeTextAutoFit(shape, props.autoFit);
    // Before any bullet: core 0.17.0, still in the peer range, sizes a bullet's hanging
    // indent from the level at that moment and never moves it afterwards.
    props.paragraphs?.forEach((paragraph, index) => {
      if (paragraph.level !== undefined) api.setParagraphLevel(shape, index, paragraph.level);
    });
    if (props.bullets !== undefined) api.setShapeBulletStyle(shape, props.bullets);
    props.paragraphs?.forEach((paragraph, index) => {
      if (paragraph.bullet !== undefined) api.setParagraphBullet(shape, index, paragraph.bullet);
    });
    if (props.paragraphSpacing) {
      const { before, after } = props.paragraphSpacing;
      const spacing = {
        ...(before !== undefined ? { beforePts: before } : {}),
        ...(after !== undefined ? { afterPts: after } : {}),
      };
      for (let i = 0; i < api.getShapeParagraphCount(shape); i++) {
        api.setParagraphSpacing(shape, i, spacing);
      }
    }
    style(shape, props);
    await visit(content.elements, { ...context, slide, shape, scope: 'shape' });
  });
}
export interface ShapeProps extends Bounds, ShapeStyle, Children {
  preset: Parameters<typeof api.addSlideShape>[1]['preset'];
  text?: string;
  format?: api.TextFormat;
  /** Alignment and anchor of `text` inside the shape. */
  align?: api.ParagraphAlignment;
  anchor?: api.TextAnchor;
}
export function Shape(props: ShapeProps): Node {
  return node('Shape', async (context) => {
    if ((props.align || props.anchor) && props.text === undefined)
      throw new Error('Shape align and anchor position its text; set text as well.');
    const shape = created(
      context,
      api.addSlideShape(requireSlide(context), {
        ...bounds(props),
        preset: props.preset,
        ...(props.name ? { name: props.name } : {}),
        ...(props.text !== undefined ? { text: props.text } : {}),
      }),
    );
    style(shape, props);
    if (props.format) api.setShapeTextFormat(shape, props.format);
    if (props.align) api.setShapeAlignment(shape, props.align);
    if (props.anchor) api.setShapeTextAnchor(shape, props.anchor);
    await visit(props.children, { ...context, shape, scope: 'shape' });
  });
}
export interface ImageProps extends Bounds, ShapeStyle, Children {
  data: Uint8Array;
  format?: api.ImageFormat;
  fit?: api.ImageFit;
}
export function Image(props: ImageProps): Node {
  return node('Image', async (context) => {
    const shape = created(
      context,
      api.addSlideImage(requireSlide(context), props.data, {
        ...bounds(props),
        ...(props.format ? { format: props.format } : {}),
        ...(props.fit ? { fit: props.fit } : {}),
        ...(props.name ? { name: props.name } : {}),
      }),
    );
    style(shape, props);
    await visit(props.children, { ...context, shape, scope: 'shape' });
  });
}
// No ShapeStyle: a clip's picture already carries the `ppaction://media` click
// action that makes it playable, which `click` would replace.
export type MediaProps = Bounds &
  Children &
  api.SlideMediaSource & {
    name?: string;
    poster?: Uint8Array;
    posterFormat?: api.ImageFormat;
  };
export function Media(props: MediaProps): Node {
  return node('Media', async (context) => {
    const { x: _x, y: _y, width: _width, height: _height, children, ...media } = props;
    const shape = created(
      context,
      api.addSlideMedia(requireSlide(context), { ...media, ...bounds(props) }),
    );
    await visit(children, { ...context, shape, scope: 'shape' });
  });
}
export interface ChartProps extends Bounds, ShapeStyle, Children {
  spec: api.ChartSpec;
}
export function Chart(props: ChartProps): Node {
  return node('Chart', async (context) => {
    const shape = created(
      context,
      api.addSlideChart(requireSlide(context), {
        ...bounds(props),
        spec: props.spec,
        ...(props.name ? { name: props.name } : {}),
      }),
    );
    style(shape, props);
    await visit(props.children, { ...context, shape, scope: 'shape' });
  });
}
export type CellBorderSide = 'left' | 'right' | 'top' | 'bottom';
export interface CellStyle {
  fill?: api.Color;
  format?: api.TextFormat;
  anchor?: api.TextAnchor;
  align?: api.ParagraphAlignment;
  /** Border per side. `width` is in points; without it the line has no explicit width. */
  borders?: Partial<Record<CellBorderSide, { color: api.Color; width?: number }>>;
}
export interface TableCellInfo {
  /** Zero-based; row 0 is the header. */
  row: number;
  column: number;
  /** A cell's text; for `paragraphs`, the run texts joined, one line per paragraph. */
  value: string;
}
/** Exactly one of `text` and `paragraphs`. An empty cell is the string `''`. */
type TableCellContent =
  | { text: string; paragraphs?: never }
  | { text?: never; paragraphs: readonly api.ParagraphSpec[] };
export type TableCellSpec = TableCellContent & {
  /** This cell is the top-left of a merge. The positions it covers stay in `rows` as `''`. */
  colSpan?: number;
  rowSpan?: number;
};
export type TableCell = string | TableCellSpec;
export interface TableProps extends Bounds, Children {
  rows: readonly (readonly TableCell[])[];
  columnWidths?: readonly number[];
  rowHeights?: readonly number[];
  cellStyle?: CellStyle;
  headerStyle?: CellStyle;
  stripeFill?: api.Color;
  /**
   * Style for one cell, merged over `cellStyle`, `headerStyle` and
   * `stripeFill` — for a status column, a total row, a highlighted figure.
   */
  styleCell?: (cell: TableCellInfo) => CellStyle | undefined;
}
function cellBorders(borders: NonNullable<CellStyle['borders']>) {
  const sides: Partial<Record<CellBorderSide, { color: api.Color; widthEmu?: number }>> = {};
  for (const side of ['left', 'right', 'top', 'bottom'] as const) {
    const border = borders[side];
    if (!border) continue;
    sides[side] = {
      color: border.color,
      ...(border.width !== undefined ? { widthEmu: api.pt(border.width) } : {}),
    };
  }
  return sides;
}
type CellPosition = { row: number; col: number };
const positionKey = (row: number, col: number) => `${row}:${col}`;
const cellSpec = (cell: TableCell): TableCellSpec =>
  typeof cell === 'string' ? { text: cell } : cell;
const cellValue = (spec: TableCellSpec): string =>
  spec.paragraphs
    ? spec.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n')
    : spec.text;

/** Maps every position a merge covers to that merge's top-left cell. */
function coveredPositions(rows: TableProps['rows']): Map<string, CellPosition> {
  const covered = new Map<string, CellPosition>();
  rows.forEach((cells, row) =>
    cells.forEach((cell, col) => {
      const { rowSpan = 1, colSpan = 1 } = cellSpec(cell);
      // Clamped to the grid: the core rejects an oversized span later, and an
      // unbounded loop must not run first.
      const lastRow = Math.min(row + rowSpan, rows.length);
      const lastCol = Math.min(col + colSpan, cells.length);
      for (let r = row; r < lastRow; r++)
        for (let c = col; c < lastCol; c++)
          if (r !== row || c !== col) covered.set(positionKey(r, c), { row, col });
    }),
  );
  return covered;
}
function checkCells(rows: TableProps['rows'], covered: Map<string, CellPosition>) {
  rows.forEach((cells, row) =>
    cells.forEach((cell, col) => {
      const spec = cellSpec(cell);
      if ((spec.text === undefined) === (spec.paragraphs === undefined))
        throw new Error(`Table cell (${row}, ${col}) accepts either text or paragraphs.`);
      const anchor = covered.get(positionKey(row, col));
      // PowerPoint paints only the merge's top-left cell, so anything written
      // here would vanish from the slide without a trace.
      if (anchor && cell !== '')
        throw new Error(
          `Table cell (${row}, ${col}) is covered by the merge at (${anchor.row}, ${anchor.col}); write it as ''.`,
        );
    }),
  );
}
export function Table(props: TableProps): Node {
  return node('Table', async (context) => {
    const covered = coveredPositions(props.rows);
    checkCells(props.rows, covered);
    const shape = created(
      context,
      api.addSlideTable(requireSlide(context), {
        ...bounds(props),
        rows: props.rows.map((cells) => cells.map((cell) => cellSpec(cell).text ?? '')),
        ...(props.columnWidths ? { colWidths: props.columnWidths.map(api.inches) } : {}),
        ...(props.rowHeights ? { rowHeights: props.rowHeights.map(api.inches) } : {}),
      }),
    );
    props.rows.forEach((row, r) =>
      row.forEach((entry, c) => {
        // A merged block is painted from its top-left cell alone, fill and all
        // four borders, so a covered position gets no style and no styleCell call.
        if (covered.has(positionKey(r, c))) return;
        const spec = cellSpec(entry);
        const cell = api.getTableCell(shape, r, c);
        const header = r === 0 ? props.headerStyle : undefined;
        const own = props.styleCell?.({ row: r, column: c, value: cellValue(spec) });
        const striped = r > 0 && r % 2 === 0 ? props.stripeFill : undefined;
        const appearance: CellStyle = {
          ...props.cellStyle,
          ...header,
          ...(striped !== undefined ? { fill: striped } : {}),
          ...own,
          format: { ...props.cellStyle?.format, ...header?.format, ...own?.format },
        };
        if (appearance.fill) api.setTableCellFill(cell, appearance.fill);
        if (spec.paragraphs) {
          // addSlideTable bakes the deck's body-text color into each run so text
          // stays readable on an inverted color map; new runs must keep it.
          const built = api.getTableCellParagraphs(cell)[0]?.elements[0];
          const base = {
            ...(built?.kind === 'r' && built.format ? writableFormat(built.format) : {}),
            ...appearance.format,
          };
          const { align } = appearance;
          api.setTableCellParagraphs(
            cell,
            spec.paragraphs.map((paragraph) => ({
              ...paragraph,
              ...(paragraph.align === undefined && align ? { align } : {}),
              runs: paragraph.runs.map((run) => ({ ...run, format: { ...base, ...run.format } })),
            })),
          );
        } else {
          if (appearance.format) api.setTableCellTextFormat(cell, appearance.format);
          if (appearance.align) api.setTableCellAlignment(cell, appearance.align);
        }
        if (appearance.anchor) api.setTableCellAnchor(cell, appearance.anchor);
        // Only when a side is set: an empty call would still give every cell a <a:tcPr>.
        const borders = cellBorders({
          ...props.cellStyle?.borders,
          ...header?.borders,
          ...own?.borders,
        });
        if (Object.keys(borders).length > 0) api.setTableCellBorders(cell, borders);
      }),
    );
    props.rows.forEach((row, r) =>
      row.forEach((entry, c) => {
        const { rowSpan = 1, colSpan = 1 } = cellSpec(entry);
        if (rowSpan !== 1 || colSpan !== 1)
          api.mergeTableCells(shape, { row: r, col: c, rowSpan, colSpan });
      }),
    );
    await visit(props.children, { ...context, shape, scope: 'shape' });
  });
}
export interface LineProps {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color?: api.Color;
  /** Stroke width in points. */
  width?: number;
  name?: string;
}
export function Line(props: LineProps): Node {
  return node('Line', (context) => {
    created(
      context,
      api.addSlideLine(requireSlide(context), {
        from: { x: api.inches(props.x1), y: api.inches(props.y1) },
        to: { x: api.inches(props.x2), y: api.inches(props.y2) },
        ...(props.color !== undefined ? { color: props.color } : {}),
        ...(props.width !== undefined ? { widthEmu: api.pt(props.width) } : {}),
        ...(props.name ? { name: props.name } : {}),
      }),
    );
  });
}
export interface GroupProps extends Children {
  name?: string;
}
/** Groups the shapes its children create, so they move and resize as one object. */
export function Group(props: GroupProps): Node {
  return node('Group', async (context) => {
    requireSlide(context);
    const members: api.SlideShapeData[] = [];
    await visit(props.children, { ...context, members });
    created(context, api.groupShapes(members, props.name ? { name: props.name } : {}));
  });
}
function selectShape(slide: api.SlideData, target: Target): api.SlideShapeData {
  const matches = api.getSlideShapes(slide).filter((shape) => {
    if ('name' in target) return api.getShapeName(shape) === target.name;
    if ('id' in target) return api.getShapeId(shape) === target.id;
    const ph = target.placeholder;
    return (
      (ph.type !== undefined || ph.idx !== undefined) &&
      (ph.type === undefined || api.getShapePlaceholderType(shape) === ph.type) &&
      (ph.idx === undefined || api.getShapePlaceholderIdx(shape) === ph.idx)
    );
  });
  if (matches.length !== 1)
    throw new Error(`Shape reference matched ${matches.length} shapes: ${JSON.stringify(target)}`);
  return matches[0]!;
}
export interface FillProps extends Children {
  target: Target;
  text?: string;
  format?: api.TextFormat;
  /** Omitted keeps whatever the template's `<a:bodyPr>` says. */
  autoFit?: api.TextAutoFit;
}
export function Fill(props: FillProps): Node {
  return node('Fill', (context) => {
    const shape = selectShape(requireSlide(context), props.target);
    api.setShapeText(shape, props.text ?? literal(props.children));
    if (props.format) api.setShapeTextFormat(shape, props.format);
    if (props.autoFit) api.setShapeTextAutoFit(shape, props.autoFit);
  });
}
export function Remove(props: { target: Target }): Node {
  return node('Remove', (context) => {
    const slide = requireSlide(context);
    api.removeShape(selectShape(slide, props.target));
  });
}
