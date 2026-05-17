// `resolveDrawingColor` — applies ECMA-376 §20.1.2.3.x color transforms
// (lumMod, lumOff, shade, tint, satMod, hueMod, alpha, gray, inv, comp)
// to the inner srgbClr / schemeClr / sysClr / prstClr element. This is
// the single resolver behind both the run-format cascade and (eventually)
// fill-format readers.

import { describe, expect, it } from 'vitest';
import { resolveDrawingColor } from '../src/api/index.ts';
import { parseXml } from '../src/internal/xml/index.ts';

const parseColorEl = (xml: string) => parseXml(xml).root;

describe('fn API: resolveDrawingColor', () => {
  it('returns srgb base color unchanged when no transforms are present', () => {
    const el = parseColorEl(
      `<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="FF0000"/>`,
    );
    expect(resolveDrawingColor(el, null)).toBe('#FF0000');
  });

  it('darkens via shade', () => {
    // shade=50000 (50%) of pure red → #800000
    const el = parseColorEl(
      `<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="FF0000">
         <a:shade val="50000"/>
       </a:srgbClr>`,
    );
    expect(resolveDrawingColor(el, null)).toBe('#800000');
  });

  it('lightens via tint', () => {
    // tint=50000 (50%) of pure red mixes 50% white → #FF8080
    const el = parseColorEl(
      `<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="FF0000">
         <a:tint val="50000"/>
       </a:srgbClr>`,
    );
    expect(resolveDrawingColor(el, null)).toBe('#FF8080');
  });

  it('combines lumMod + lumOff (PowerPoint "Lighter 60%" preset)', () => {
    // PowerPoint's "Accent 1, Lighter 60%" theme variant uses
    // lumMod=40000 + lumOff=60000 — the canonical recipe per §20.1.2.3.20-21.
    const el = parseColorEl(
      `<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="4472C4">
         <a:lumMod val="40000"/>
         <a:lumOff val="60000"/>
       </a:srgbClr>`,
    );
    const hex = resolveDrawingColor(el, null);
    expect(hex).toMatch(/^#[0-9A-F]{6}$/);
    // The result should be lighter than the input (higher luminance).
    expect(hex).not.toBe('#4472C4');
  });

  it('resolves schemeClr when a theme is provided', () => {
    const el = parseColorEl(
      `<a:schemeClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="accent1"/>`,
    );
    const theme = {
      name: 't',
      dark1: '#000000',
      light1: '#FFFFFF',
      dark2: '#222222',
      light2: '#EEEEEE',
      accent1: '#4472C4',
      accent2: '',
      accent3: '',
      accent4: '',
      accent5: '',
      accent6: '',
      hyperlink: '',
      followedHyperlink: '',
    };
    expect(resolveDrawingColor(el, theme)).toBe('#4472C4');
  });

  it('returns null for schemeClr without a theme', () => {
    const el = parseColorEl(
      `<a:schemeClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="accent1"/>`,
    );
    expect(resolveDrawingColor(el, null)).toBeNull();
  });

  it('inverts via inv', () => {
    const el = parseColorEl(
      `<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="334466">
         <a:inv/>
       </a:srgbClr>`,
    );
    expect(resolveDrawingColor(el, null)).toBe('#CCBB99');
  });

  it('grays via gray', () => {
    const el = parseColorEl(
      `<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="FF0000">
         <a:gray/>
       </a:srgbClr>`,
    );
    const hex = resolveDrawingColor(el, null);
    // Grayscale of pure red ≈ #4D4D4D (0.3 luminance weight).
    expect(hex).toMatch(/^#[0-9A-F]{6}$/);
    expect(hex!.slice(1, 3)).toBe(hex!.slice(3, 5));
    expect(hex!.slice(3, 5)).toBe(hex!.slice(5, 7));
  });

  it('reads sysClr lastClr', () => {
    const el = parseColorEl(
      `<a:sysClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="windowText" lastClr="123456"/>`,
    );
    expect(resolveDrawingColor(el, null)).toBe('#123456');
  });
});
