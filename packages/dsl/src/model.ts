import * as api from '@office-kit/pptx';

export type Child = Node | string | number | boolean | null | undefined | readonly Child[];
export interface Children {
  children?: Child;
}
export interface RawContext {
  presentation: api.PresentationData;
  slide?: api.SlideData;
  shape?: api.SlideShapeData;
}
/** Where `jsxDEV` saw an element. Line and column are 1-based, as TypeScript and esbuild emit them. */
export interface JsxSource {
  fileName: string;
  lineNumber: number;
  columnNumber: number;
}
export interface Context extends RawContext {
  mode: 'compose' | 'edit';
  /** Sources of the elements being evaluated, outermost first. */
  sources: readonly JsxSource[];
  /** Every created shape's `sources`, by slide part name and shape id. */
  shapeSources: Map<string, readonly JsxSource[]>;
  scope: 'presentation' | 'slide' | 'shape';
  originals: readonly api.SlideData[];
  deferred: Array<() => void | Promise<void>>;
  /** Set inside a `Group`: every shape created there, in drawing order. */
  members?: api.SlideShapeData[];
}
export interface Node {
  readonly kind: string;
  readonly source?: JsxSource;
  readonly evaluate: (context: Context) => void | Promise<void>;
}
export interface PresentationNode extends Node {
  readonly compile: () => Promise<api.PresentationData>;
}
export function node(kind: string, evaluate: Node['evaluate']): Node {
  return { kind, evaluate };
}
export async function visit(children: Child, context: Context): Promise<void> {
  if (children == null || typeof children === 'boolean') return;
  if (typeof children === 'string' || typeof children === 'number') {
    if (String(children).trim()) throw new Error('Text content must be inside Text or Fill.');
    return;
  }
  if ('evaluate' in children) {
    const { source } = children;
    try {
      await children.evaluate(
        source ? { ...context, sources: [...context.sources, source] } : context,
      );
    } catch (cause) {
      if (!source) throw cause;
      const { fileName, lineNumber, columnNumber } = source;
      throw new Error(
        `${fileName}:${lineNumber}:${columnNumber} <${children.kind}>: ${cause instanceof Error ? cause.message : String(cause)}`,
        { cause },
      );
    }
  } else {
    for (const child of children) await visit(child, context);
  }
}
export function requireSlide(context: Context): api.SlideData {
  if (!context.slide || context.scope !== 'slide')
    throw new Error('This element must be a child of Slide.');
  return context.slide;
}
/** Reports a new shape to the enclosing `Group`, if any, and records the JSX that made it. */
export function created(context: Context, shape: api.SlideShapeData): api.SlideShapeData {
  if (context.sources.length > 0)
    context.shapeSources.set(sourceKey(api.getShapeSlide(shape), shape), context.sources);
  context.members?.push(shape);
  return shape;
}

// The core replaces slide and shape handles whenever a slide or shape list
// changes (removing the source slides, grouping), and `getGroupChildren`
// makes new ones on every call. Part names and shape ids survive all of that,
// so compile records by them and binds the result to the final slides.
const sourceKey = (slide: api.SlideData, shape: api.SlideShapeData) =>
  `${api.getSlidePartName(slide)}#${api.getShapeId(shape)}`;
const jsxSources = new WeakMap<api.SlideData, Map<number, readonly JsxSource[]>>();

export function bindJsxSources(
  presentation: api.PresentationData,
  recorded: Context['shapeSources'],
): void {
  if (recorded.size === 0) return;
  for (const slide of api.getSlides(presentation)) {
    const byId = new Map<number, readonly JsxSource[]>();
    for (const shape of api.getSlideShapes(slide)) {
      const sources = recorded.get(sourceKey(slide, shape));
      if (sources) byId.set(api.getShapeId(shape), sources);
    }
    jsxSources.set(slide, byId);
  }
}

/**
 * The JSX elements, outermost first, that were being evaluated when `compile`
 * created this shape. `undefined` without the development JSX runtime, and for
 * shapes a template, `Slide from` or `Raw` supplied.
 */
export function getShapeJsxSources(shape: api.SlideShapeData): readonly JsxSource[] | undefined {
  return jsxSources.get(api.getShapeSlide(shape))?.get(api.getShapeId(shape));
}
export function literal(children: Child): string {
  if (children == null || typeof children === 'boolean') return '';
  if (typeof children === 'string' || typeof children === 'number') return String(children);
  if ('evaluate' in children)
    throw new Error('Use paragraphs for rich text; element children are not text.');
  return children.map(literal).join('');
}

/** Separate inline text from shape-scoped escape hatches without evaluating them. */
export function textContent(children: Child): { text: string; elements: Node[] } {
  const elements: Node[] = [];
  function collect(child: Child): string {
    if (child == null || typeof child === 'boolean') return '';
    if (typeof child === 'string' || typeof child === 'number') return String(child);
    if ('evaluate' in child) {
      elements.push(child);
      return '';
    }
    return child.map(collect).join('');
  }
  return { text: collect(children), elements };
}
