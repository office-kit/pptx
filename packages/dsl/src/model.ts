import type * as api from '@office-kit/pptx';

export type Child = Node | string | number | boolean | null | undefined | readonly Child[];
export interface Children {
  children?: Child;
}
export interface RawContext {
  presentation: api.PresentationData;
  slide?: api.SlideData;
  shape?: api.SlideShapeData;
}
export interface Context extends RawContext {
  mode: 'compose' | 'edit';
  scope: 'presentation' | 'slide' | 'shape';
  originals: readonly api.SlideData[];
  deferred: Array<() => void | Promise<void>>;
}
export interface Node {
  readonly kind: string;
  readonly source?: { fileName: string; lineNumber: number; columnNumber: number };
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
    try {
      await children.evaluate(context);
    } catch (cause) {
      if (!children.source) throw cause;
      const { fileName, lineNumber, columnNumber } = children.source;
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
