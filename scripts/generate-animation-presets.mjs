// Generates src/internal/presentationml/animation-presets.generated.ts from the
// `<p:timing>` the reference desktop app (Mac 16.113) writes for every effect of its Entrance,
// Emphasis and Exit galleries, checked in under
// test/fixtures/native/animations/ (provenance in test/fixtures/SOURCES.md).
//
//   node scripts/generate-animation-presets.mjs
//   pnpm exec oxfmt src/internal/presentationml/animation-presets.generated.ts
//
// Each capture is one effect on shape 2. What is kept of it is the effect node:
// its preset numbers, the attributes it carries besides them, whether it joins a
// build group, and everything under it after its start condition — the
// behaviours, with `p:` dropped, the cTn ids left out (the builder numbers them)
// and the target written `{T}`. Where an effect takes options (a fly's
// direction, a filter's pattern) the values those options decide are written as
// `{…}` placeholders the builder fills in.
//
// test/fn-animation-native.test.ts builds every effect again at its defaults and
// compares it with the capture byte for byte.
import { readFileSync, readdirSync, writeFileSync } from 'node:fs';

const ROOT = new URL('../test/fixtures/native/animations/', import.meta.url);
const TARGET = new URL(
  '../src/internal/presentationml/animation-presets.generated.ts',
  import.meta.url,
);

// Capture number -> the `AnimationEffect` token it is.
const TOKENS = {
  1: 'appear',
  2: 'blindsIn',
  3: 'checkerboardIn',
  4: 'dissolveIn',
  5: 'flyIn',
  6: 'peekIn',
  7: 'randomBarsIn',
  8: 'shapeIn',
  9: 'splitIn',
  10: 'stripsIn',
  11: 'wedgeIn',
  12: 'wheelIn',
  13: 'wipeIn',
  14: 'expandIn',
  15: 'fadeIn',
  16: 'swivelIn',
  17: 'zoomIn',
  18: 'centerRevolveIn',
  19: 'floatIn',
  20: 'growTurnIn',
  21: 'riseUpIn',
  22: 'spinnerIn',
  23: 'basicZoomIn',
  24: 'stretchIn',
  25: 'boomerangIn',
  26: 'bounceIn',
  27: 'creditsIn',
  28: 'curveUpIn',
  29: 'dropIn',
  30: 'flipIn',
  31: 'floatingIn',
  32: 'pinwheelIn',
  33: 'spiralIn',
  34: 'basicSwivelIn',
  35: 'whipIn',
  36: 'blindsOut',
  37: 'checkerboardOut',
  38: 'disappear',
  39: 'dissolveOut',
  40: 'flyOut',
  41: 'peekOut',
  42: 'randomBarsOut',
  43: 'shapeOut',
  44: 'splitOut',
  45: 'stripsOut',
  46: 'wedgeOut',
  47: 'wheelOut',
  48: 'wipeOut',
  49: 'contractOut',
  50: 'fadeOut',
  51: 'swivelOut',
  52: 'zoomOut',
  53: 'centerRevolveOut',
  54: 'collapseOut',
  55: 'floatOut',
  56: 'shrinkTurnOut',
  57: 'sinkDownOut',
  58: 'spinnerOut',
  59: 'basicZoomOut',
  60: 'stretchyOut',
  61: 'boomerangOut',
  62: 'bounceOut',
  63: 'creditsOut',
  64: 'curveDownOut',
  65: 'dropOut',
  66: 'flipOut',
  67: 'floatingOut',
  68: 'pinwheelOut',
  69: 'spiralOut',
  70: 'basicSwivelOut',
  71: 'whipOut',
  72: 'fillColor',
  73: 'fontColor',
  74: 'growShrink',
  75: 'lineColor',
  76: 'spin',
  77: 'transparency',
  78: 'boldFlash',
  79: 'brushColor',
  80: 'complementaryColor',
  81: 'complementaryColor2',
  82: 'contrastingColor',
  83: 'darken',
  84: 'desaturate',
  85: 'lighten',
  86: 'objectColor',
  87: 'pulse',
  88: 'underline',
  89: 'colorPulse',
  90: 'growWithColor',
  91: 'shimmer',
  92: 'teeter',
  93: 'blink',
  94: 'boldReveal',
  95: 'wave',
};

// The families whose `<p:animEffect filter>` an option decides.
const FILTER_OPTION_FAMILIES = new Set([
  'blinds',
  'checkerboard',
  'randomBars',
  'shape',
  'split',
  'strips',
  'wheel',
  'wipe',
  'peek',
]);

const TARGET_XML = '<tgtEl><spTgt spid="2"/></tgtEl>';

/** The value of every `<p:tav>` under the `<p:anim>` that animates `attrName`, in order. */
const placeholdTavs = (body, attrName, names) => {
  const re = new RegExp(
    `(<anim [^>]*>(?:(?!</anim>).)*<attrName>${attrName}</attrName>(?:(?!</anim>).)*</anim>)`,
  );
  const match = re.exec(body);
  if (match === null) throw new Error(`no ${attrName} animation`);
  let at = 0;
  const replaced = match[1].replace(/Val val="[^"]*"/g, (whole) => {
    const name = names[at++];
    if (name === undefined) throw new Error(`more keyframes than expected for ${attrName}`);
    return whole.replace(/val="[^"]*"/, `val="{${name}}"`);
  });
  if (at !== names.length) throw new Error(`fewer keyframes than expected for ${attrName}`);
  return body.replace(match[1], replaced);
};

// The markup every body repeats, written shorter. The builder expands them in
// the reverse order (`expandAnimationTemplate`), so the table is part of the
// generated file rather than duplicated there.
const ABBREVIATIONS = [
  '<anim calcmode="lin" valueType="num">',
  '</anim>',
  '<cBhvr>',
  '</cBhvr>',
  '<tavLst>',
  '</tavLst>',
  ' fill="hold"',
  '<cTn ',
  '</cTn>',
  '<animEffect transition="',
  '</animEffect>',
  '<set>',
  '</set>',
  '<animRot by="',
  '</animRot>',
  '<animClr clrSpc="',
  '</animClr>',
  '<animScale>',
  '</animScale>',
  ' dur="',
  '"/>',
];
for (const marker of ['~', '[', '|', '!', '@', '^']) {
  for (const file of ['entrance', 'exit', 'emphasis']) {
    for (const name of readdirSync(new URL(`${file}/`, ROOT))) {
      const xml = readFileSync(new URL(`${file}/${name}`, ROOT), 'utf8').replace(
        /^<!--.*-->\n/,
        '',
      );
      if (xml.includes(marker)) throw new Error(`${name} uses ${marker}, which abbreviations need`);
    }
  }
}

// Keyframes, attribute names and start delays take the most room, so they get
// shapes of their own: `[tm|string]`, `[tm!number]`, `@name@` and `^delay^`.
const abbreviate = (body) => {
  let out = body
    .replaceAll(/<tav tm="(\d+)"><val><strVal val="([^"]*)"\/><\/val><\/tav>/g, '[$1|$2]')
    .replaceAll(/<tav tm="(\d+)"><val><fltVal val="([^"]*)"\/><\/val><\/tav>/g, '[$1!$2]')
    .replaceAll(/<attrNameLst><attrName>([^<]*)<\/attrName><\/attrNameLst>/g, '@$1@')
    .replaceAll(/<stCondLst><cond delay="(\w+)"\/><\/stCondLst>/g, '^$1^');
  ABBREVIATIONS.forEach((long, at) => {
    out = out.replaceAll(long, `~${String.fromCharCode(97 + at)}`);
  });
  return out;
};

const presets = [];
for (const group of ['entrance', 'exit', 'emphasis']) {
  for (const file of readdirSync(new URL(`${group}/`, ROOT)).sort()) {
    const number = Number.parseInt(file, 10);
    const token = TOKENS[number];
    if (token === undefined) throw new Error(`no token for ${file}`);
    const xml = readFileSync(new URL(`${group}/${file}`, ROOT), 'utf8');
    const effect =
      /<p:par><p:cTn id="5" presetID="(\d+)" presetClass="(\w+)" presetSubtype="(\d+)"([^>]*?)( grpId="0")? nodeType="clickEffect"><p:stCondLst><p:cond delay="0"\/><\/p:stCondLst>(.*?)<\/p:cTn><\/p:par><\/p:childTnLst><\/p:cTn><\/p:par><\/p:childTnLst><\/p:cTn><\/p:par><\/p:childTnLst><\/p:cTn><p:prevCondLst>/.exec(
        xml,
      );
    if (effect === null) throw new Error(`no single effect in ${file}`);
    const [, presetId, presetClass, presetSubtype, attrs, grouped, inner] = effect;
    const bldP = /<p:bldLst>(.*)<\/p:bldLst>/.exec(xml)?.[1] ?? null;
    // Which build entry the reference desktop app writes: one animating the shape's background
    // with its text, one for the text alone, or none at all.
    const build =
      bldP === null
        ? 'none'
        : bldP === '<p:bldP spid="2" grpId="0" animBg="1"/>'
          ? 'shape'
          : bldP === '<p:bldP spid="2" grpId="0"/>'
            ? 'text'
            : null;
    if (build === null) throw new Error(`unexpected build entry in ${file}: ${bldP}`);
    if ((build === 'none') !== (grouped === undefined)) {
      throw new Error(`${file}: grpId without a build entry, or the other way round`);
    }
    let body = inner
      .replaceAll(/<(\/?)p:/g, '<$1')
      .replaceAll(/<cTn id="\d+"/g, '<cTn')
      .replaceAll(TARGET_XML, '{T}');
    if (body.includes('spid=')) throw new Error(`${file}: a target other than the shape`);
    const family = token.replace(/(In|Out)$/, '');
    if (FILTER_OPTION_FAMILIES.has(family)) body = body.replace(/filter="[^"]*"/, 'filter="{F}"');
    if (family === 'fly') {
      body = placeholdTavs(body, 'ppt_x', ['X0', 'X1']);
      body = placeholdTavs(body, 'ppt_y', ['Y0', 'Y1']);
    }
    if (family === 'peek') {
      body = placeholdTavs(body, 'ppt_y', ['P0', 'P1']).replace(
        '<attrName>ppt_y</attrName>',
        '<attrName>{A}</attrName>',
      );
    }
    presets.push({
      token,
      presetClass,
      presetId: Number(presetId),
      presetSubtype: Number(presetSubtype),
      attrs: attrs.trim(),
      build,
      body: abbreviate(body),
    });
  }
}
if (presets.length !== 95) throw new Error(`expected 95 effects, found ${presets.length}`);

const out = [
  '// GENERATED by scripts/generate-animation-presets.mjs from',
  '// test/fixtures/native/animations/ — the `<p:timing>` the reference desktop app (Mac 16.113) writes',
  '// for each effect of its Entrance, Emphasis and Exit galleries; provenance in',
  '// test/fixtures/SOURCES.md. Do not edit.',
  '',
  '/**',
  ' * One gallery effect as the reference desktop app writes it:',
  ' * `[presetClass, presetID, presetSubtype, extra effect-node attributes, build',
  ' * entry, body]`. The body is what follows the effect node’s start condition,',
  ' * with `p:` and the cTn ids left out and the target written `{T}`.',
  ' */',
  'export type AnimationPresetTemplate = readonly [',
  "  presetClass: 'entr' | 'exit' | 'emph',",
  '  presetId: number,',
  '  presetSubtype: number,',
  '  attrs: string,',
  "  build: 'shape' | 'text' | 'none',",
  '  body: string,',
  '];',
  '',
  '/** Markup the bodies write as `~a`, `~b`, … — see `expandAnimationTemplate`. */',
  `export const ANIMATION_TEMPLATE_ABBREVIATIONS: readonly string[] = ${JSON.stringify(ABBREVIATIONS)};`,
  '',
  '/** Every effect of the three galleries, in gallery order (entrance, exit, emphasis). */',
  'export const ANIMATION_PRESET_TEMPLATES = {',
  ...presets.map(
    (p) =>
      `  ${p.token}: [${JSON.stringify(p.presetClass)}, ${p.presetId}, ${p.presetSubtype}, ${JSON.stringify(p.attrs)}, ${JSON.stringify(p.build)}, ${JSON.stringify(p.body)}],`,
  ),
  '} as const satisfies Record<string, AnimationPresetTemplate>;',
  '',
];
writeFileSync(TARGET, out.join('\n'));
