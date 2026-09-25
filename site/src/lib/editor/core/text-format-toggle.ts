import type { TextFormat } from '@office-kit/pptx';

export type TextFormatToggle =
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strike'
  | 'superscript'
  | 'subscript';

export function textFormatActive(
  formats: readonly TextFormat[],
  property: TextFormatToggle,
): boolean {
  return (
    formats.length > 0 &&
    formats.every((format) => {
      if (property === 'superscript') return (format.baseline ?? 0) > 0;
      if (property === 'subscript') return (format.baseline ?? 0) < 0;
      const value = format[property];
      return (
        value === true || (typeof value === 'string' && value !== 'none' && value !== 'noStrike')
      );
    })
  );
}

export function toggleTextFormat(
  formats: readonly TextFormat[],
  property: TextFormatToggle,
): TextFormat {
  const active = textFormatActive(formats, property);
  if (property === 'superscript' || property === 'subscript')
    return { baseline: active ? 0 : property === 'superscript' ? 0.3 : -0.25 };
  return { [property]: !active };
}
