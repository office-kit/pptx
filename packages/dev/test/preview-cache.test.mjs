import { test } from 'node:test';
import assert from 'node:assert/strict';
import { compile, Presentation, Slide, Text } from '@office-kit/pptx-dsl';
import { loadPresentation, savePresentation } from '@office-kit/pptx';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { renderPreview } from '../src/preview-cache.ts';

async function deck(title = 'Original') {
  return savePresentation(
    await compile(
      Presentation({
        children: [title, 'Unchanged'].map((text) =>
          Slide({ children: Text({ x: 1, y: 1, width: 4, height: 1, children: text }) }),
        ),
      }),
    ),
  );
}
async function render(bytes, cache) {
  return renderPreview(await loadPresentation(bytes), bytes, cache);
}

test('a one-slide edit reuses the other SVG and text after a fresh evaluation', async () => {
  const before = await render(await deck());
  const after = await render(await deck('Edited'), before.cache);
  const names = Object.keys(before.cache.slides);
  assert.notEqual(after.cache.slides[names[0]], before.cache.slides[names[0]]);
  assert.equal(after.cache.slides[names[1]], before.cache.slides[names[1]]);
  assert.match(after.slides[0], /Edited/);
  assert.equal(after.slideTexts[0], 'Edited');
  const repeated = await render(await deck('Edited'), after.cache);
  assert.equal(repeated.cache.slides[names[0]], after.cache.slides[names[0]]);
});

test('shared themes, relationships and media invalidate every cached slide', async () => {
  const bytes = await deck();
  const before = await render(bytes);
  for (const resource of ['theme', 'relationships', 'media']) {
    const parts = unzipSync(bytes);
    if (resource === 'media') {
      parts['ppt/media/test.png'] = new Uint8Array([1, 2, 3]);
      parts['[Content_Types].xml'] = strToU8(
        strFromU8(parts['[Content_Types].xml']).replace(
          '</Types>',
          '<Default Extension="png" ContentType="image/png"/></Types>',
        ),
      );
    } else {
      const name = Object.keys(parts).find((name) =>
        resource === 'theme'
          ? /^ppt\/theme\/.*\.xml$/.test(name)
          : /presentation.xml.rels$/.test(name),
      );
      assert.ok(name);
      parts[name] = strToU8(strFromU8(parts[name]) + '\n<!-- modified -->');
    }
    const after = await render(zipSync(parts), before.cache);
    for (const name of Object.keys(before.cache.slides)) {
      assert.notEqual(after.cache.slides[name], before.cache.slides[name], resource);
    }
  }
});
