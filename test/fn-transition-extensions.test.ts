// Extension transitions ([MS-PPTX] §2.2.1): the p14 effects, the
// p15 preset transitions and p159 Morph, written as the reference desktop app writes them — an
// mc:AlternateContent whose choice requires the extension and whose fallback
// is a <p:fade/>.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { INTERNAL_PACKAGE, SLIDE_PART_NAME } from '../src/api/_internal-symbols.ts';
import {
  type SlideTransition,
  type TransitionOptions,
  getSlideTransition,
  getSlideTransitionSound,
  getSlideXmlString,
  getSlides,
  loadPresentation,
  savePresentation,
  setSlideTransition,
  setSlideTransitionSound,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = fileURLToPath(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url));
const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

const NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const NS_P15 = 'http://schemas.microsoft.com/office/powerpoint/2012/main';
const NS_P159 = 'http://schemas.microsoft.com/office/powerpoint/2015/09/main';

const firstSlide = async () => {
  const pres = await loadPresentation(await readFile(fixture));
  return { pres, slide: getSlides(pres)[0]! };
};

// Every extension effect, with the options the reference desktop app's Effect Options menus
// write, the element expected in the choice and the namespace it requires.
const CASES: ReadonlyArray<readonly [TransitionOptions, string, 'p14' | 'p15' | 'p159']> = [
  [{ effect: 'morph' }, '<p159:morph option="byObject"/>', 'p159'],
  [{ effect: 'morph', morphOption: 'byChar' }, '<p159:morph option="byChar"/>', 'p159'],
  [
    { effect: 'reveal', direction: 'r', thruBlack: true },
    '<p14:reveal dir="r" thruBlk="1"/>',
    'p14',
  ],
  [{ effect: 'flash' }, '<p14:flash/>', 'p14'],
  [{ effect: 'prstTrans', preset: 'fallOver' }, '<p15:prstTrans prst="fallOver"/>', 'p15'],
  [
    { effect: 'prstTrans', preset: 'pageCurlSingle', invertX: true },
    '<p15:prstTrans prst="pageCurlSingle" invX="1"/>',
    'p15',
  ],
  [
    { effect: 'prstTrans', preset: 'curtains', invertY: true },
    '<p15:prstTrans prst="curtains" invY="1"/>',
    'p15',
  ],
  [{ effect: 'ripple', direction: 'center' }, '<p14:ripple dir="center"/>', 'p14'],
  [{ effect: 'ripple', direction: 'ld' }, '<p14:ripple dir="ld"/>', 'p14'],
  [{ effect: 'honeycomb' }, '<p14:honeycomb/>', 'p14'],
  [
    { effect: 'glitter', direction: 'u', pattern: 'hexagon' },
    '<p14:glitter dir="u" pattern="hexagon"/>',
    'p14',
  ],
  [{ effect: 'vortex', direction: 'r' }, '<p14:vortex dir="r"/>', 'p14'],
  [
    { effect: 'shred', direction: 'out', pattern: 'rectangle' },
    '<p14:shred dir="out" pattern="rectangle"/>',
    'p14',
  ],
  [{ effect: 'switch', direction: 'l' }, '<p14:switch dir="l"/>', 'p14'],
  [{ effect: 'flip', direction: 'r' }, '<p14:flip dir="r"/>', 'p14'],
  [{ effect: 'gallery', direction: 'l' }, '<p14:gallery dir="l"/>', 'p14'],
  [{ effect: 'prism', direction: 'u' }, '<p14:prism dir="u"/>', 'p14'],
  [{ effect: 'doors', direction: 'horz' }, '<p14:doors dir="horz"/>', 'p14'],
  [{ effect: 'prism', isInverted: true }, '<p14:prism isInverted="1"/>', 'p14'],
  [{ effect: 'warp', direction: 'out' }, '<p14:warp dir="out"/>', 'p14'],
  [{ effect: 'pan', direction: 'd' }, '<p14:pan dir="d"/>', 'p14'],
  [{ effect: 'ferris', direction: 'r' }, '<p14:ferris dir="r"/>', 'p14'],
  [{ effect: 'conveyor', direction: 'l' }, '<p14:conveyor dir="l"/>', 'p14'],
  [
    { effect: 'prism', direction: 'd', isContent: true },
    '<p14:prism dir="d" isContent="1"/>',
    'p14',
  ],
  [{ effect: 'window', direction: 'vert' }, '<p14:window dir="vert"/>', 'p14'],
  [
    { effect: 'prism', isContent: true, isInverted: true },
    '<p14:prism isContent="1" isInverted="1"/>',
    'p14',
  ],
  [
    { effect: 'flythrough', direction: 'out', hasBounce: true },
    '<p14:flythrough dir="out" hasBounce="1"/>',
    'p14',
  ],
  [{ effect: 'wheelReverse', spokes: 1 }, '<p14:wheelReverse spokes="1"/>', 'p14'],
];

// The effect element of the choice, made a document of its own so it can be
// validated against [MS-PPTX]'s schema fragment.
const extensionElement = (xml: string): string => {
  const match = /<mc:Choice\b[^>]*>(<p:transition\b[^>]*>)(<(p14|p15|p159):[^>]*\/>)/.exec(xml);
  if (!match) throw new Error(`no extension effect in ${xml}`);
  const prefix = match[3]!;
  const ns = { p14: NS_P14, p15: NS_P15, p159: NS_P159 }[prefix as 'p14'];
  return match[2]!.replace(/^<(\S+?)(\/?>| )/, `<$1 xmlns:${prefix}="${ns}"$2`);
};

// What a reader that only knows ECMA-376 sees: the fallback in place of the
// AlternateContent.
const resolveFallback = (xml: string): string =>
  xml.replace(
    /<mc:AlternateContent\b[^>]*>[\s\S]*?<mc:Fallback>([\s\S]*?)<\/mc:Fallback><\/mc:AlternateContent>/g,
    '$1',
  );

// What a p14-aware reader sees, less the two things ECMA-376's schema cannot
// know: the extension effect element and the p14:dur attribute.
const resolveChoiceForEcma = (xml: string): string =>
  xml.replace(
    /<mc:AlternateContent\b[^>]*><mc:Choice\b[^>]*>([\s\S]*?)<\/mc:Choice>[\s\S]*?<\/mc:AlternateContent>/g,
    (_, choice: string) =>
      choice.replace(/ p14:dur="\d+"/, '').replace(/<(p14|p15|p159):[^>]*\/>/, ''),
  );

describe('The reference desktop app’s extension transitions', () => {
  it.each(CASES)(
    'writes %j in an AlternateContent with a fade fallback',
    async (options, element, required) => {
      const { pres, slide } = await firstSlide();
      setSlideTransition(slide, { ...options, durationMs: 1600 });
      const xml = getSlideXmlString(slide);
      expect(xml).toContain(`<mc:Choice xmlns:${required}=`);
      expect(xml).toContain(`Requires="${required}"`);
      expect(xml).toContain(`<p:transition spd="slow" p14:dur="1600">${element}</p:transition>`);
      expect(xml).toContain(
        '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback>',
      );
      const expected = { ...options, speed: 'slow', durationMs: 1600 };
      expect(getSlideTransition(slide)).toEqual(
        options.effect === 'morph' ? { morphOption: 'byObject', ...expected } : expected,
      );
      const reloaded = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
      const read = getSlideTransition(reloaded)!;
      expect(read).toEqual(getSlideTransition(slide));
      // Writing back what was read changes nothing: the read model is lossless.
      setSlideTransition(reloaded, read as TransitionOptions);
      expect(getSlideXmlString(reloaded)).toBe(xml);
    },
  );

  skipIfNoXmllint(
    'keeps every choice and fallback schema-valid',
    async () => {
      for (const [options] of CASES) {
        const { slide } = await firstSlide();
        setSlideTransition(slide, { ...options, durationMs: 1250, advanceAfterMs: 3000 });
        setSlideTransitionSound(slide, { kind: 'stop' });
        const xml = getSlideXmlString(slide);
        expectSchemaValid(resolveFallback(xml), 'pml');
        expectSchemaValid(resolveChoiceForEcma(xml), 'pml');
        expectSchemaValid(extensionElement(xml), 'pptxTransitions');
      }
      // Three xmllint runs per case.
    },
    60_000,
  );

  it('declares p14 beside a p15 or p159 choice only when a duration needs it', async () => {
    const { slide } = await firstSlide();
    setSlideTransition(slide, { effect: 'prstTrans', preset: 'origami' });
    let xml = getSlideXmlString(slide);
    expect(xml).toMatch(
      /<mc:Choice xmlns:p15="[^"]+" Requires="p15"><p:transition><p15:prstTrans prst="origami"\/>/,
    );
    expect(xml).not.toContain('p14:dur');
    setSlideTransition(slide, { effect: 'morph', durationMs: 2000 });
    xml = getSlideXmlString(slide);
    expect(xml).toContain(`xmlns:p159="${NS_P159}"`);
    expect(xml).toContain(`xmlns:p14="${NS_P14}"`);
    expect(xml).toContain('Requires="p159"');
    expect(xml.match(/<p:transition\b/g)).toHaveLength(2);
  });

  it('keeps a transition sound in both the choice and the fallback', async () => {
    const { pres, slide } = await firstSlide();
    setSlideTransition(slide, { effect: 'vortex', direction: 'l', durationMs: 3000 });
    setSlideTransitionSound(slide, { kind: 'stop' });
    setSlideTransition(slide, { effect: 'honeycomb', durationMs: 3000 });
    expect(getSlideXmlString(slide).match(/<p:sndAc><p:endSnd\/><\/p:sndAc>/g)).toHaveLength(2);
    const reloaded = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
    expect(getSlideTransitionSound(reloaded)).toEqual({ kind: 'stop' });
    expect(getSlideTransition(reloaded)?.effect).toBe('honeycomb');
  });

  it.each<[TransitionOptions, RegExp]>([
    [{ effect: 'prstTrans' }, /needs a preset/],
    [{ effect: 'prstTrans', preset: 'spin' as never }, /preset/],
    [{ effect: 'glitter', pattern: 'strip' }, /pattern/],
    [{ effect: 'shred', pattern: 'hexagon' }, /pattern/],
    [{ effect: 'switch', direction: 'u' }, /direction "u" is not valid for effect "switch"/],
    [{ effect: 'ripple', direction: 'l' }, /direction "l" is not valid for effect "ripple"/],
    [{ effect: 'doors', direction: 'in' }, /direction "in" is not valid for effect "doors"/],
    [{ effect: 'morph', morphOption: 'byLine' as never }, /morphOption/],
  ])('rejects %j without touching the slide', async (options, error) => {
    const { slide } = await firstSlide();
    const before = getSlideXmlString(slide);
    expect(() => setSlideTransition(slide, options)).toThrow(error);
    expect(getSlideXmlString(slide)).toBe(before);
  });

  it('ignores options on effects whose element has no such attribute', async () => {
    const { slide } = await firstSlide();
    setSlideTransition(slide, {
      effect: 'honeycomb',
      direction: 'l',
      pattern: 'hexagon',
      isContent: true,
      hasBounce: true,
      thruBlack: true,
      spokes: 3,
    });
    expect(getSlideXmlString(slide)).toContain('<p14:honeycomb/>');
    expect(getSlideTransition(slide)).toEqual({ effect: 'honeycomb' });
  });

  it('reports an effect element of an unknown namespace by its prefixed name and keeps it', async () => {
    const { pres, slide } = await firstSlide();
    setSlideTransition(slide, { effect: 'flash' });
    const part = pres[INTERNAL_PACKAGE].getPart(slide[SLIDE_PART_NAME])!;
    part.data = new TextEncoder().encode(
      new TextDecoder()
        .decode(part.data)
        .replace('<p14:flash/>', '<p99:sparkle xmlns:p99="urn:example:p99"/>')
        .replace('Requires="p14"', 'xmlns:p99="urn:example:p99" Requires="p99"'),
    );
    const reloaded = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
    expect(getSlideTransition(reloaded)).toEqual<SlideTransition>({ effect: 'p99:sparkle' });
    const again = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
    expect(getSlideXmlString(again)).toContain('<p99:sparkle');
  });

  it('does not mistake a p14 element for the ECMA-376 one of the same name', async () => {
    const { pres, slide } = await firstSlide();
    setSlideTransition(slide, { effect: 'flash' });
    const part = pres[INTERNAL_PACKAGE].getPart(slide[SLIDE_PART_NAME])!;
    part.data = new TextEncoder().encode(
      new TextDecoder().decode(part.data).replace('<p14:flash/>', '<p14:fade/>'),
    );
    const reloaded = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
    expect(getSlideTransition(reloaded)?.effect).toBe('p14:fade');
  });
});
