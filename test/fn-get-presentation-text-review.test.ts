// `getPresentationTextReview(pres, options)` — headline track, open
// questions and wording flags, the output of the skill's deck-text script.

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getPresentationTextReview,
  inches,
  setShapeTextFormat,
  setSlideNotes,
  type PresentationData,
  type SlideData,
} from '../src/api/index.ts';

const addText = (
  slide: SlideData,
  text: string,
  options: { name?: string; size?: number; font?: string } = {},
): void => {
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(8),
    h: inches(1),
    text,
    ...(options.name !== undefined ? { name: options.name } : {}),
  });
  setShapeTextFormat(shape, {
    ...(options.size !== undefined ? { size: options.size } : {}),
    ...(options.font !== undefined ? { font: options.font } : {}),
  });
};

const deck = (): PresentationData => {
  const pres = createPresentation();
  const cover = addBlankSlide(pres);
  addText(cover, '権限はどこで決まるの？', { name: 'Headline' });
  addText(cover, '行レベルの権限が効く');
  addText(cover, 'limit = [要確認: 上限]', { font: 'Courier New' });
  const guessed = addBlankSlide(pres);
  addText(guessed, '目次', { size: 40 });
  addText(guessed, 'DB 制約で守り切る', { size: 12 });
  setSlideNotes(guessed, 'ここで RLS が効いてくる\n[要確認: 移行日]');
  const document = addBlankSlide(pres);
  addText(document, 'この 3 つの論点です', { name: 'Headline' });
  return pres;
};

const HEADLINES = `# Headline track
p01  権限はどこで決まるの？
p02? 目次
p03  この 3 つの論点です

"?" marks a guessed headline (largest text). Name the headline Text "Headline".`;

const FLAGS = `
# [要確認] left on slides (fill each or move it to the notes before delivery)
p01 limit = [要確認: 上限]

# [要確認] in notes (report each to the user)
p02 notes [要確認: 移行日]

# Wording flags (judge each against references/writing-ja.md)
p01 [効く] 行レベルの権限が効く
p02 [intensifying 切る] DB 制約で守り切る
p02 notes [効く] ここで RLS が効いてくる`;

describe('fn API: getPresentationTextReview', () => {
  it('lists headlines (guessing unnamed ones), open questions and wording flags', () => {
    expect(getPresentationTextReview(deck())).toBe(`${HEADLINES}\n${FLAGS}`);
  });

  it('adds every slide text, code and notes with full', () => {
    expect(getPresentationTextReview(deck(), { full: true })).toBe(`${HEADLINES}

# Slide text

## p01 権限はどこで決まるの？
- 行レベルの権限が効く
- [code] limit = [要確認: 上限]

## p02 目次
- DB 制約で守り切る
  notes: ここで RLS が効いてくる [要確認: 移行日]

## p03 この 3 つの論点です
${FLAGS}`);
  });

  it('adds the headline mix in talk mode', () => {
    expect(getPresentationTextReview(deck(), { mode: 'talk' })).toBe(`${HEADLINES}

# Talk headline mix (hand-made reference range in brackets)
！ 0% [7-20%]
？ 33% [10-15%]
polite form 33% [10-12%]
shortest 2 chars (p02) [3-6]; median 11
${FLAGS}`);
  });

  it('adds headline checks past the cover in document mode', () => {
    expect(getPresentationTextReview(deck(), { mode: 'document' })).toBe(`${HEADLINES}

# Document headline checks
p03 polite ending (use plain form); refers to another slide; element count in headline (state the content, or check it matches the body): この 3 つの論点です
${FLAGS}`);
  });
});
