import type { ShapeParagraphElement } from '@office-kit/pptx';

type TextElement<T> = Exclude<T, { readonly kind: 'math' }>;

/**
 * A paragraph's runs, fields and breaks without its equations. The core
 * text-range APIs count only these, so editor offsets must skip equations too.
 */
export const textElements = <T extends ShapeParagraphElement>(
  elements: ReadonlyArray<T>,
): TextElement<T>[] => elements.filter((e): e is TextElement<T> => e.kind !== 'math');
