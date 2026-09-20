import * as api from '@office-kit/pptx';
import {
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
    background?: string;
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
  fill?: string | api.GradientFillOptions;
  stroke?: { color: string; width?: number } | false;
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
export interface TextProps extends Bounds, ShapeStyle, api.TextFormat, Children {
  paragraphs?: readonly api.ParagraphSpec[];
  align?: api.ParagraphAlignment;
  anchor?: api.TextAnchor;
  autoFit?: api.TextAutoFit;
}
export function Text(props: TextProps): Node {
  return node('Text', async (context) => {
    const slide = requireSlide(context);
    const content = textContent(props.children);
    if (props.paragraphs && content.text.trim())
      throw new Error('Text accepts either text children or paragraphs.');
    const shape = api.addSlideTextBox(slide, {
      ...bounds(props),
      ...(props.name ? { name: props.name } : {}),
      text: content.text,
    });
    if (props.paragraphs) api.setShapeParagraphs(shape, props.paragraphs);
    api.setShapeTextFormat(shape, props);
    if (props.align) api.setShapeAlignment(shape, props.align);
    if (props.anchor) api.setShapeTextAnchor(shape, props.anchor);
    if (props.autoFit) api.setShapeTextAutoFit(shape, props.autoFit);
    style(shape, props);
    await visit(content.elements, { ...context, slide, shape, scope: 'shape' });
  });
}
export interface ShapeProps extends Bounds, ShapeStyle, Children {
  preset: Parameters<typeof api.addSlideShape>[1]['preset'];
  text?: string;
  format?: api.TextFormat;
}
export function Shape(props: ShapeProps): Node {
  return node('Shape', async (context) => {
    const shape = api.addSlideShape(requireSlide(context), {
      ...bounds(props),
      preset: props.preset,
      ...(props.name ? { name: props.name } : {}),
      ...(props.text !== undefined ? { text: props.text } : {}),
    });
    style(shape, props);
    if (props.format) api.setShapeTextFormat(shape, props.format);
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
    const shape = api.addSlideImage(requireSlide(context), props.data, {
      ...bounds(props),
      ...(props.format ? { format: props.format } : {}),
      ...(props.fit ? { fit: props.fit } : {}),
      ...(props.name ? { name: props.name } : {}),
    });
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
    const shape = api.addSlideMedia(requireSlide(context), { ...media, ...bounds(props) });
    await visit(children, { ...context, shape, scope: 'shape' });
  });
}
export interface ChartProps extends Bounds, ShapeStyle, Children {
  spec: api.ChartSpec;
}
export function Chart(props: ChartProps): Node {
  return node('Chart', async (context) => {
    const shape = api.addSlideChart(requireSlide(context), {
      ...bounds(props),
      spec: props.spec,
      ...(props.name ? { name: props.name } : {}),
    });
    style(shape, props);
    await visit(props.children, { ...context, shape, scope: 'shape' });
  });
}
export interface CellStyle {
  fill?: string;
  format?: api.TextFormat;
  anchor?: api.TextAnchor;
}
export interface TableProps extends Bounds, Children {
  rows: readonly (readonly string[])[];
  columnWidths?: readonly number[];
  rowHeights?: readonly number[];
  cellStyle?: CellStyle;
  headerStyle?: CellStyle;
  stripeFill?: string;
}
export function Table(props: TableProps): Node {
  return node('Table', async (context) => {
    const shape = api.addSlideTable(requireSlide(context), {
      ...bounds(props),
      rows: props.rows,
      ...(props.columnWidths ? { colWidths: props.columnWidths.map(api.inches) } : {}),
      ...(props.rowHeights ? { rowHeights: props.rowHeights.map(api.inches) } : {}),
    });
    props.rows.forEach((row, r) =>
      row.forEach((_, c) => {
        const cell = api.getTableCell(shape, r, c);
        const appearance = {
          ...props.cellStyle,
          ...(r === 0 ? props.headerStyle : {}),
          format: { ...props.cellStyle?.format, ...(r === 0 ? props.headerStyle?.format : {}) },
        };
        if (appearance.fill) api.setTableCellFill(cell, appearance.fill);
        if (r > 0 && r % 2 === 0 && props.stripeFill) api.setTableCellFill(cell, props.stripeFill);
        if (appearance.format) api.setTableCellTextFormat(cell, appearance.format);
        if (appearance.anchor) api.setTableCellAnchor(cell, appearance.anchor);
      }),
    );
    await visit(props.children, { ...context, shape, scope: 'shape' });
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
}
export function Fill(props: FillProps): Node {
  return node('Fill', (context) => {
    const shape = selectShape(requireSlide(context), props.target);
    api.setShapeText(shape, props.text ?? literal(props.children));
    if (props.format) api.setShapeTextFormat(shape, props.format);
  });
}
export function Remove(props: { target: Target }): Node {
  return node('Remove', (context) => {
    const slide = requireSlide(context);
    api.removeShape(selectShape(slide, props.target));
  });
}
