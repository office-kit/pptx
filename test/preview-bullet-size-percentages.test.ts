import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setParagraphBullet,
  setShapeRunFormat,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

describe('preview percentage bullet sizes', () => {
  it.each([0.5, 1.25, 4])('sizes a %sx marker relative to the first run', async (ratio) => {
    const pres = createPresentation();
    const shape = addSlideTextBox(addBlankSlide(pres), {
      x: inches(1),
      y: inches(1),
      w: inches(5),
      h: inches(3),
      text: 'Large text',
    });
    setParagraphBullet(shape, 0, 'bullet');
    setShapeRunFormat(shape, 0, 0, { size: 30 });
    const zip = readZip(await savePresentation(pres));
    const entry = zip.entries.find((part) => part.name === 'ppt/slides/slide1.xml')!;
    const xml = new TextDecoder().decode(entry.data);
    expect(xml).toContain('<a:buChar');
    const data = new TextEncoder().encode(
      xml.replace('<a:buChar', `<a:buSzPct val="${ratio * 100000}"/><a:buChar`),
    );
    const imported = await loadPresentation(
      writeZip(zip.entries.map((part) => (part === entry ? { ...part, data } : part))),
    );
    const html = renderSlideToSvg(imported, getSlides(imported)[0]!, {
      textLayout: 'foreignObject',
    });
    const marker = html.match(/<span style="([^"]*)">•<\/span>/);
    expect(marker).not.toBeNull();
    // Percent CSS would inherit the host application's font size instead of
    // the paragraph's 30pt first run. Emit an absolute size in slide units.
    expect(marker![1]).toContain(`font-size:${(((30 * 4) / 3) * ratio).toFixed(2)}px`);
  });
});
