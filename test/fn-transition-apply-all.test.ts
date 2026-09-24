import { expect, it } from 'vitest';
import * as pptx from '../src/api/index.ts';
import { INTERNAL_PACKAGE, SLIDE_DOCUMENT, SLIDE_PART_NAME } from '../src/api/_internal-symbols.ts';
import { partName, resolveTarget } from '../src/internal/opc/index.ts';
import { NS, parseXml } from '../src/internal/xml/index.ts';

it('copies complete transitions and remaps sound relationships without duplication', async () => {
  const p = pptx.createPresentation();
  const source = pptx.addBlankSlide(p);
  const destination = pptx.addBlankSlide(p);
  const pkg = p[INTERNAL_PACKAGE];
  pkg.addPart(partName('/ppt/media/transition.wav'), 'audio/x-wav', new Uint8Array([1, 2, 3]));
  const rels = pkg.getRels(source[SLIDE_PART_NAME])!;
  rels.items.push({
    id: 'rId99',
    type: NS.officeDocRels + '/audio',
    target: '../media/transition.wav',
    targetMode: 'Internal',
  });
  pkg.setRels(source[SLIDE_PART_NAME], rels);
  source[SLIDE_DOCUMENT].root.children.push(
    parseXml(
      `<p:transition xmlns:p="${NS.pml}" xmlns:r="${NS.officeDocRels}" spd="slow" advClick="0" advTm="1250"><p:fade thruBlk="1"/><p:sndAc><p:stSnd><p:snd r:embed="rId99" name="Chime"/></p:stSnd></p:sndAc><p:extLst/></p:transition>`,
    ).root,
  );
  // Commit the raw source before saving/reloading it.
  pptx.setSlideAdvanceTiming(source, { advanceOnClick: false, advanceAfterMs: 1250 });
  pptx.setSlideTransition(destination, { effect: 'push', direction: 'l' });
  pptx.applySlideTransitionToAll(p, source);
  expect(pptx.getSlideTransition(destination)).toEqual(pptx.getSlideTransition(source));
  const copiedRels = pkg.getRels(destination[SLIDE_PART_NAME])!;
  const sound = copiedRels.items.find((r) => r.type.endsWith('/audio'))!;
  expect(sound.id).not.toBe('rId99');
  expect(resolveTarget(destination[SLIDE_PART_NAME], sound.target)).toBe(
    '/ppt/media/transition.wav',
  );
  pptx.applySlideTransitionToAll(p, source);
  expect(pkg.getRels(destination[SLIDE_PART_NAME])).toEqual(copiedRels);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  const slides = pptx.getSlides(loaded);
  expect(pptx.getSlideTransition(slides[1]!)).toEqual(pptx.getSlideTransition(slides[0]!));
  const xml = new TextDecoder().decode(
    loaded[INTERNAL_PACKAGE].getPart(slides[1]![SLIDE_PART_NAME])!.data,
  );
  expect(xml).toContain('name="Chime"');
  expect(xml).toContain('p:extLst');
  expect(xml).toContain(`r:embed="${sound.id}"`);
  pptx.clearSlideTransition(slides[0]!);
  pptx.applySlideTransitionToAll(loaded, slides[0]!);
  expect(pptx.getSlideTransition(slides[1]!)).toBeNull();
});

it('rejects broken relationships and foreign sources before changing the package', async () => {
  const p = pptx.createPresentation();
  const source = pptx.addBlankSlide(p);
  pptx.addBlankSlide(p);
  pptx.addBlankSlide(p);
  source[SLIDE_DOCUMENT].root.children.push(
    parseXml(
      `<p:transition xmlns:p="${NS.pml}" xmlns:r="${NS.officeDocRels}"><p:sndAc><p:stSnd><p:snd r:embed="missing"/></p:stSnd></p:sndAc></p:transition>`,
    ).root,
  );
  pptx.setSlideAdvanceTiming(source, { advanceOnClick: true, advanceAfterMs: null });
  const before = await pptx.savePresentation(p);
  expect(() => pptx.applySlideTransitionToAll(p, source)).toThrow(/relationship is missing/);
  expect(await pptx.savePresentation(p)).toEqual(before);
  const other = pptx.createPresentation();
  expect(() => pptx.applySlideTransitionToAll(p, pptx.addBlankSlide(other))).toThrow(
    /does not belong/,
  );
  expect(await pptx.savePresentation(p)).toEqual(before);
});

it('changes only effects while retaining timing, sound and extensions, and validates atomically', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  const original = parseXml(
    `<p:transition xmlns:p="${NS.pml}" spd="slow" advClick="0" advTm="2500"><p:fade/><p:sndAc><p:endSnd/></p:sndAc><p:extLst/></p:transition>`,
  ).root;
  slide[SLIDE_DOCUMENT].root.children.push(original);
  const retained = structuredClone(original.children.slice(1));
  pptx.setSlideTransitionEffect(slide, { effect: 'split', direction: 'out', orientation: 'vert' });
  expect(pptx.getSlideTransition(slide)).toEqual({
    effect: 'split',
    direction: 'out',
    orientation: 'vert',
    speed: 'slow',
    advanceOnClick: false,
    advanceAfterMs: 2500,
  });
  const transition = slide[SLIDE_DOCUMENT].root.children.find(
    (c) => c.kind === 'element' && c.name.localName === 'transition',
  );
  expect(transition?.kind === 'element' && transition.children.slice(1)).toEqual(retained);
  const before = await pptx.savePresentation(p);
  expect(() =>
    pptx.setSlideTransitionEffect(slide, { effect: 'push', direction: 'invalid' }),
  ).toThrow();
  expect(await pptx.savePresentation(p)).toEqual(before);
  pptx.setSlideTransitionEffect(slide, { effect: 'none' });
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  expect(pptx.getSlideTransition(pptx.getSlides(loaded)[0]!)).toEqual({
    effect: 'none',
    speed: 'slow',
    advanceOnClick: false,
    advanceAfterMs: 2500,
  });
  const xml = new TextDecoder().decode(
    loaded[INTERNAL_PACKAGE].getPart(pptx.getSlides(loaded)[0]![SLIDE_PART_NAME])!.data,
  );
  expect(xml).toContain('p:endSnd');
  expect(xml).toContain('p:extLst');
});

it('reads and edits PowerPoint Choice/Fallback transitions and copies their full container', async () => {
  const p = pptx.createPresentation();
  const source = pptx.addBlankSlide(p);
  const destination = pptx.addBlankSlide(p);
  const wrapper = parseXml(
    `<mc:AlternateContent xmlns:mc="${NS.mc}" xmlns:p="${NS.pml}" xmlns:p14="${NS.p14}"><mc:Choice Requires="p14"><p:transition p14:dur="2000" advTm="3000"><p:fade/><p:sndAc><p:endSnd/></p:sndAc></p:transition></mc:Choice><mc:Fallback><p:transition advTm="3000"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>`,
  ).root;
  source[SLIDE_DOCUMENT].root.children.push(wrapper);
  expect(pptx.getSlideTransition(source)).toEqual({
    effect: 'fade',
    durationMs: 2000,
    advanceAfterMs: 3000,
  });
  pptx.setSlideAdvanceTiming(source, { advanceOnClick: false, advanceAfterMs: 4500 });
  pptx.setSlideTransitionEffect(source, { effect: 'push', direction: 'r' });
  pptx.setSlideTransitionDuration(source, 1750);
  const expected = {
    effect: 'push',
    direction: 'r',
    durationMs: 1750,
    advanceAfterMs: 4500,
    advanceOnClick: false,
  };
  expect(pptx.getSlideTransition(source)).toEqual(expected);
  pptx.applySlideTransitionToAll(p, source);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  for (const slide of pptx.getSlides(loaded)) {
    expect(pptx.getSlideTransition(slide)).toEqual(expected);
    const xml = new TextDecoder().decode(
      loaded[INTERNAL_PACKAGE].getPart(slide[SLIDE_PART_NAME])!.data,
    );
    expect(xml).toContain('Requires="p14"');
    expect(xml.match(/<p:push/g)).toHaveLength(2);
    expect(xml.match(/advTm="4500"/g)).toHaveLength(2);
    expect(xml).toContain('p:endSnd');
    pptx.clearSlideTransition(slide);
    expect(pptx.getSlideTransition(slide)).toBeNull();
    expect(
      slide[SLIDE_DOCUMENT].root.children.some(
        (c) => c.kind === 'element' && c.name.localName === 'AlternateContent',
      ),
    ).toBe(false);
  }
  pptx.setSlideTransition(destination, { effect: 'wipe', durationMs: 100 });
  expect(pptx.getSlideTransition(destination)?.durationMs).toBe(100);
});

it('duration writes round-trip with namespaces and reject invalid values without mutation', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  pptx.setSlideTransition(slide, { effect: 'fade', advanceAfterMs: 3000 });
  pptx.setSlideTransitionDuration(slide, 0);
  let loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  expect(pptx.getSlideTransition(pptx.getSlides(loaded)[0]!)?.durationMs).toBe(0);
  for (const duration of [-1, 0.1, NaN, Infinity, 4294967296]) {
    const before = await pptx.savePresentation(p);
    expect(() => pptx.setSlideTransitionDuration(slide, duration)).toThrow();
    expect(await pptx.savePresentation(p)).toEqual(before);
  }
  pptx.setSlideTransitionDuration(slide, 1234);
  loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  expect(pptx.getSlideTransition(pptx.getSlides(loaded)[0]!)).toEqual({
    effect: 'fade',
    advanceAfterMs: 3000,
    durationMs: 1234,
  });
});

it('round-trips wheel spokes and preserves timing through effect edits', async () => {
  const p = pptx.createPresentation();
  const source = pptx.addBlankSlide(p);
  const destination = pptx.addBlankSlide(p);
  pptx.setSlideTransition(source, {
    effect: 'wheel',
    spokes: 3,
    durationMs: 1200,
    advanceAfterMs: 2500,
  });
  expect(pptx.getSlideTransition(source)?.spokes).toBe(3);
  for (const spokes of [-1, 1.5, NaN, Infinity, 4294967296]) {
    expect(() => pptx.setSlideTransitionEffect(source, { effect: 'wheel', spokes })).toThrow();
    expect(pptx.getSlideTransition(source)?.spokes).toBe(3);
  }
  pptx.setSlideTransitionEffect(source, { effect: 'wheel', spokes: 8 });
  pptx.applySlideTransitionToAll(p, source);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  for (const slide of pptx.getSlides(loaded))
    expect(pptx.getSlideTransition(slide)).toEqual({
      effect: 'wheel',
      spokes: 8,
      durationMs: 1200,
      advanceAfterMs: 2500,
    });
  pptx.setSlideTransitionEffect(destination, { effect: 'fade', spokes: 8 });
  expect(pptx.getSlideTransition(destination)?.spokes).toBeUndefined();
});

it('selects usable compatibility branches by namespace and retains unsupported effects', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  const destination = pptx.addBlankSlide(p);
  slide[SLIDE_DOCUMENT].root.prefixDecls.set('modern', NS.p14);
  slide[SLIDE_DOCUMENT].root.children.push(
    parseXml(
      `<mc:AlternateContent xmlns:mc="${NS.mc}" xmlns:p="${NS.pml}" xmlns:p14="${NS.p14}" xmlns:future="urn:future">
      <mc:Choice Requires="modern future"><p:transition advTm="111"><p:push/></p:transition></mc:Choice>
      <mc:Choice Requires="modern"><p:transition p14:dur="1500"><p14:ripple dir="ld"/><p:sndAc><p:endSnd/></p:sndAc></p:transition></mc:Choice>
      <mc:Fallback><p:transition advTm="3000"><p:fade thruBlk="1"/></p:transition></mc:Fallback>
    </mc:AlternateContent>`,
    ).root,
  );
  expect(pptx.getSlideTransition(slide)).toEqual({
    effect: 'fade',
    thruBlack: true,
    advanceAfterMs: 3000,
  });
  expect(pptx.getSlideTransitionSound(slide)).toEqual({ kind: 'none' });
  pptx.setSlideTransitionDuration(slide, 1700);
  pptx.applySlideTransitionToAll(p, slide);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  for (const copy of pptx.getSlides(loaded)) {
    expect(pptx.getSlideTransition(copy)).toEqual({
      effect: 'fade',
      thruBlack: true,
      advanceAfterMs: 3000,
      durationMs: 1700,
    });
    const xml = new TextDecoder().decode(
      loaded[INTERNAL_PACKAGE].getPart(copy[SLIDE_PART_NAME])!.data,
    );
    expect(xml).toContain('ripple');
    expect(xml).toContain('urn:future');
  }
  pptx.setSlideTransitionEffect(slide, { effect: 'wipe', direction: 'u' });
  expect(pptx.getSlideTransition(slide)).toEqual({
    effect: 'wipe',
    direction: 'u',
    durationMs: 1700,
  });
  expect(pptx.getSlideTransitionSound(slide)).toEqual({ kind: 'stop' });
  expect(pptx.getSlideTransition(destination)?.effect).toBe('fade');
});

it('uses the first supported Choice and omits unsupported branches without a fallback', () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  const wrapper =
    parseXml(`<mc:AlternateContent xmlns:mc="${NS.mc}" xmlns:p="${NS.pml}" xmlns:q="${NS.p14}">
    <mc:Choice Requires="missing"><p:transition><p:push/></p:transition></mc:Choice>
    <mc:Choice Requires="q"><p:transition q:dur="900"><p:split dir="in"/><p:sndAc><p:endSnd/></p:sndAc></p:transition></mc:Choice>
    <mc:Choice Requires="q"><p:transition><p:cut/></p:transition></mc:Choice>
    <mc:Fallback><p:transition><p:fade/></p:transition></mc:Fallback>
  </mc:AlternateContent>`).root;
  slide[SLIDE_DOCUMENT].root.children.push(wrapper);
  expect(pptx.getSlideTransition(slide)).toEqual({
    effect: 'split',
    direction: 'in',
    durationMs: 900,
  });
  expect(pptx.getSlideTransitionSound(slide)).toEqual({ kind: 'stop' });
  wrapper.children = wrapper.children
    .filter((child) => child.kind === 'element' && child.name.localName === 'Choice')
    .slice(0, 1);
  expect(pptx.getSlideTransition(slide)).toBeNull();
  expect(pptx.getSlideTransitionSound(slide)).toEqual({ kind: 'none' });
});

it('round-trips reverse wheel with a legacy fallback and preserves timing and sound', async () => {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  pptx.addBlankSlide(p);
  pptx.setSlideTransition(slide, {
    effect: 'fade',
    durationMs: 1750,
    advanceAfterMs: 2500,
    advanceOnClick: false,
  });
  pptx.setSlideTransitionSound(slide, { kind: 'stop' });
  pptx.setSlideTransitionEffect(slide, { effect: 'wheelReverse', spokes: 3 });
  expect(pptx.getSlideTransition(slide)).toMatchObject({
    effect: 'wheelReverse',
    spokes: 3,
    durationMs: 1750,
    advanceAfterMs: 2500,
    advanceOnClick: false,
  });
  expect(pptx.getSlideTransitionSound(slide)).toEqual({ kind: 'stop' });
  const before = await pptx.savePresentation(p);
  expect(() =>
    pptx.setSlideTransitionEffect(slide, { effect: 'wheelReverse', spokes: -1 }),
  ).toThrow();
  expect(await pptx.savePresentation(p)).toEqual(before);
  pptx.applySlideTransitionToAll(p, slide);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  for (const item of pptx.getSlides(loaded)) {
    expect(pptx.getSlideTransition(item)).toEqual(pptx.getSlideTransition(slide));
    expect(pptx.getSlideTransitionSound(item)).toEqual({ kind: 'stop' });
    const xml = new TextDecoder().decode(
      loaded[INTERNAL_PACKAGE].getPart(item[SLIDE_PART_NAME])!.data,
    );
    expect(xml).toContain('Requires="p14"');
    expect(xml).toContain('p14:wheelReverse');
    expect(xml).toMatch(/<p:wheel\s[^>]*spokes="3"/);
    pptx.setSlideTransitionEffect(item, { effect: 'fade' });
    expect(pptx.getSlideTransition(item)?.effect).toBe('fade');
    expect(pptx.getSlideTransitionSound(item)).toEqual({ kind: 'stop' });
  }
  pptx.setSlideTransition(slide, { effect: 'wheelReverse', spokes: 8 });
  expect(pptx.getSlideTransition(slide)).toMatchObject({ effect: 'wheelReverse', spokes: 8 });
});
