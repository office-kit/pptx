// CSS for a Text Art preset's "A" swatch, previewed from the deck theme. Used
// by the Text Art gallery and the Shape Format tab's in-ribbon Text Art strip.
import {
  asColor,
  type Color,
  type ColorTransform,
  type PatternPreset,
  type PresentationTheme,
  type TextFormat,
} from '@office-kit/pptx';
import { resolveColor } from './theme-color.ts';
import type { TextArtPreset } from './text-art-presets.ts';

// Swatch scale: one point of the preset draws one CSS pixel.
const EMU_PER_PX = 12700;
const px = (emu = 0) => `${emu / EMU_PER_PX}px`;
const rgba = (hex: string, opacity = 1) => {
  const [r, g, b] = [1, 3, 5].map((at) => parseInt(hex.slice(at, at + 2), 16));
  return `rgba(${r}, ${g}, ${b}, ${opacity})`;
};
// CSS stand-ins for the preset patterns the gallery uses.
const PATTERNS: Partial<Record<PatternPreset, (fg: string, bg: string) => string>> = {
  dkUpDiag: (fg, bg) => `repeating-linear-gradient(-45deg, ${fg} 0 2px, ${bg} 2px 4px)`,
  ltDnDiag: (fg, bg) => `repeating-linear-gradient(45deg, ${fg} 0 1px, ${bg} 1px 4px)`,
  narHorz: (fg, bg) => `repeating-linear-gradient(0deg, ${fg} 0 1px, ${bg} 1px 3px)`,
  pct50: (fg, bg) => `repeating-conic-gradient(${fg} 0 25%, ${bg} 0 50%) 0 0 / 2px 2px`,
};

export function textArtSwatchStyle(preset: TextArtPreset, theme: PresentationTheme | null): string {
  const paint = (color: Color, transforms: readonly ColorTransform[] = []) =>
    resolveColor(color, transforms, theme) ?? '#000000';
  const fill = (format: TextFormat): string => {
    const textFill = format.textFill;
    if (textFill?.kind === 'gradient') {
      // DrawingML measures from 3 o'clock, CSS from 12.
      const stops = textFill.stops.map(
        (stop) => `${paint(stop.color, stop.colorTransforms)} ${stop.offset * 100}%`,
      );
      return `linear-gradient(${(textFill.angleDeg ?? 90) + 90}deg, ${stops.join(', ')})`;
    }
    if (textFill?.kind === 'pattern') {
      const fg = paint(asColor(textFill.foreground) ?? 'tx1', textFill.foregroundTransforms);
      const bg = paint(asColor(textFill.background) ?? 'bg1', textFill.backgroundTransforms);
      return PATTERNS[textFill.preset]?.(fg, bg) ?? fg;
    }
    // A preset without a fill keeps the inherited text color, Text 1 here.
    return format.color ? paint(format.color, format.colorTransforms) : paint('tx1');
  };
  const format = preset.format;
  const filters: string[] = [];
  if (format.shadow) {
    const {
      offsetEmu = 0,
      angleDeg = 45,
      blurEmu = 0,
      color = '#000000',
      colorTransforms,
      opacity,
    } = format.shadow;
    const radians = (angleDeg * Math.PI) / 180;
    filters.push(
      `drop-shadow(${px(offsetEmu * Math.cos(radians))} ${px(offsetEmu * Math.sin(radians))} ${px(blurEmu)} ${rgba(paint(color, colorTransforms), opacity)})`,
    );
  }
  if (format.glow) {
    const glow = rgba(paint(format.glow.color, format.glow.colorTransforms), format.glow.opacity);
    filters.push(
      `drop-shadow(0 0 ${px(format.glow.radiusEmu)} ${glow})`,
      `drop-shadow(0 0 ${px(format.glow.radiusEmu)} ${glow})`,
    );
  }
  // CSS has no inner shadow for glyphs; those swatches show fill and outline only.
  const outline =
    format.outline?.color && format.outline.widthEmu
      ? `${px(format.outline.widthEmu)} ${paint(format.outline.color, format.outline.colorTransforms)}`
      : '0';
  return [
    // The shorthand resets the clip, so the clip follows it here.
    `background: ${fill(format)}`,
    'background-clip: text',
    '-webkit-background-clip: text',
    `font-weight: ${format.bold ? 700 : 400}`,
    `-webkit-text-stroke: ${outline}`,
    `filter: ${filters.join(' ') || 'none'}`,
    format.reflection
      ? `-webkit-box-reflect: below 0 linear-gradient(transparent 55%, ${rgba('#000000', format.reflection.startOpacity)})`
      : '',
  ]
    .filter(Boolean)
    .join('; ');
}
