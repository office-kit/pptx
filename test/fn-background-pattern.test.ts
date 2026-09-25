import { readFile } from 'node:fs/promises';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { describe, it, expect } from 'vitest';
import {
  createPresentation,
  addBlankSlide,
  setSlideBackgroundPatternFill,
  setSlideBackground,
  getSlideBackgroundPatternFill,
  savePresentation,
  loadPresentation,
  getSlides,
  getSlideXmlString,
} from '../src/api/index.ts';

describe('background pattern editing', () => {
  it('preserves unspecified pattern colors through editing and save/reload', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    setSlideBackgroundPatternFill(slide, {
      preset: 'pct5',
      foreground: '#123456',
      background: '#ABCDEF',
    });
    setSlideBackgroundPatternFill(slide, { preset: 'wave' });
    const expected = { preset: 'wave', foreground: '#123456', background: '#ABCDEF' };
    expect(getSlideBackgroundPatternFill(pres, slide)).toEqual(expected);
    const loaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideBackgroundPatternFill(loaded, getSlides(loaded)[0]!)).toEqual(expected);
  });
  it('retains imported theme colors and transforms when changing only the pattern', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    setSlideBackgroundPatternFill(slide, {});
    const parts = unzipSync(await savePresentation(pres));
    const path = 'ppt/slides/slide1.xml';
    parts[path] = strToU8(
      strFromU8(parts[path]!).replace(
        '<a:schemeClr val="accent1"/>',
        '<a:schemeClr val="accent1"><a:lumMod val="75000"/><a:alpha val="80000"/></a:schemeClr>',
      ),
    );
    const imported = await loadPresentation(zipSync(parts));
    const target = getSlides(imported)[0]!;
    const original = getSlideXmlString(target);
    setSlideBackgroundPatternFill(target, { preset: 'wave' });
    const expected = original.replace('prst="pct5"', 'prst="wave"');
    expect(getSlideXmlString(target)).toBe(expected);
    const reloaded = await loadPresentation(await savePresentation(imported));
    expect(getSlideXmlString(getSlides(reloaded)[0]!)).toBe(expected);
  });
  it.each(['slideLayouts', 'slideMasters'])(
    'edits patterns inherited from %s without flattening colors',
    async (source) => {
      const parts = unzipSync(
        await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)),
      );
      const pattern =
        '<a:pattFill prst="wave"><a:fgClr><a:schemeClr val="accent2"><a:lumMod val="75000"/></a:schemeClr></a:fgClr><a:bgClr><a:schemeClr val="bg2"/></a:bgClr></a:pattFill>';
      for (const [name, bytes] of Object.entries(parts)) {
        if (!/^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name)) continue;
        let xml = strFromU8(bytes).replace(/<p:bg\b[^>]*>[\s\S]*?<\/p:bg>/g, '');
        if (name.startsWith(`ppt/${source}/`))
          xml = xml.replace(/(<p:cSld\b[^>]*>)/, '$1<p:bg><p:bgPr>' + pattern + '</p:bgPr></p:bg>');
        parts[name] = strToU8(xml);
      }
      const pres = await loadPresentation(zipSync(parts));
      const slide = getSlides(pres)[0]!;
      setSlideBackgroundPatternFill(slide, { preset: 'pct20' });
      expect(getSlideXmlString(slide)).toContain(pattern.replace('prst="wave"', 'prst="pct20"'));
      expect(getSlideBackgroundPatternFill(pres, slide, { preserveTheme: true })?.background).toBe(
        'bg2',
      );
      const saved = unzipSync(await savePresentation(pres));
      for (const name of Object.keys(parts).filter(
        (name) => name.startsWith(`ppt/${source}/`) && name.endsWith('.xml'),
      )) {
        expect(strFromU8(saved[name]!)).toBe(strFromU8(parts[name]!));
      }
      setSlideBackground(slide, '#123456');
      setSlideBackgroundPatternFill(slide, {});
      expect(getSlideBackgroundPatternFill(pres, slide, { preserveTheme: true })).toEqual({
        preset: 'pct5',
        foreground: 'accent1',
        background: 'bg1',
      });
    },
  );
  it('uses native defaults and rejects invalid edits atomically', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    setSlideBackgroundPatternFill(slide, {});
    const previous = getSlideBackgroundPatternFill(pres, slide);
    expect(previous?.preset).toBe('pct5');
    expect(getSlideXmlString(slide)).toContain('<a:fgClr><a:schemeClr val="accent1"/></a:fgClr>');
    expect(getSlideXmlString(slide)).toContain('<a:bgClr><a:schemeClr val="bg1"/></a:bgClr>');
    expect(() => setSlideBackgroundPatternFill(slide, { foreground: 'invalid-color' })).toThrow();
    expect(getSlideBackgroundPatternFill(pres, slide)).toEqual(previous);
  });
});
