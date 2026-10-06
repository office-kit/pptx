import type { TextFormat } from '@office-kit/pptx';

/** Applies a typing-format patch while preserving the exclusive glyph-fill choice. */
export function mergeTextFormat(base: TextFormat | undefined, patch: TextFormat): TextFormat {
  const result = { ...base, ...patch };
  if (patch.textFill !== undefined) delete result.color;
  else if (patch.color !== undefined) delete result.textFill;
  return result;
}
