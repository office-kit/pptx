// Layer-0 quality gate: tree-shakeability.
//
// The public API is free functions. Consumers should be able to pull
// just the entries they use without dragging the rest of the API into
// the bundle. These tests bundle minimal entries and assert that
// feature-specific identifiers from unused capabilities are dropped.

import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';
import { build } from 'esbuild';
import { describe, expect, it } from 'vitest';

const REPO_ROOT = fileURLToPath(new URL('../', import.meta.url));

interface BundleResult {
  bytes: number;
  text: string;
}

const bundleEntry = async (source: string): Promise<BundleResult> => {
  const dir = mkdtempSync(join(tmpdir(), 'pptx-kit-ts-'));
  try {
    const entry = join(dir, 'entry.mjs');
    const out = join(dir, 'bundle.mjs');
    writeFileSync(entry, source, 'utf8');
    await build({
      entryPoints: [entry],
      bundle: true,
      format: 'esm',
      platform: 'browser',
      target: 'es2022',
      outfile: out,
      treeShaking: true,
      minify: false,
      alias: { 'pptx-kit': join(REPO_ROOT, 'src/index.ts') },
      logLevel: 'silent',
    });
    const text = readFileSync(out, 'utf8');
    return { bytes: Buffer.byteLength(text, 'utf8'), text };
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
};

describe('tree-shake: minimal load+save entry', () => {
  it('bundles successfully (free-function API)', async () => {
    const result = await bundleEntry(`
      import { loadPresentation, savePresentation } from 'pptx-kit';
      export async function run(bytes) {
        const pres = await loadPresentation(bytes);
        return await savePresentation(pres);
      }
    `);
    expect(result.bytes).toBeGreaterThan(0);
    expect(result.text).toContain('zip');
  });

  it("free-function bundle drops authoring identifiers it doesn't use", async () => {
    const fn = await bundleEntry(`
      import { loadPresentation, savePresentation } from 'pptx-kit';
      export async function run(bytes) {
        const pres = await loadPresentation(bytes);
        return await savePresentation(pres);
      }
    `);
    const dropped = [
      'addTable',
      'setTransition',
      'addLine',
      'addTextBox',
      'addShape',
      'duplicateSlide',
      'setHyperlink',
      'setBullets',
    ];
    for (const name of dropped) {
      expect(fn.text, `bundle still references ${name}`).not.toContain(name);
    }
  });

  it('slide+shape read free-function bundle still drops mutation identifiers', async () => {
    const fn = await bundleEntry(`
      import {
        loadPresentation,
        savePresentation,
        getSlides,
        getSlideText,
        getSlideShapes,
        getSlideLayout,
        findSlidePlaceholder,
        getShapeKind,
        getShapeId,
        getShapeName,
        getShapeText,
        getShapePosition,
        getShapeSize,
        getShapeRotation,
        getShapeFlip,
        replaceTokensInSlide,
      } from 'pptx-kit';
      export async function run(bytes) {
        const pres = await loadPresentation(bytes);
        for (const s of getSlides(pres)) {
          getSlideText(s); getSlideLayout(s);
          const ph = findSlidePlaceholder(s, 'title');
          for (const sh of getSlideShapes(s)) {
            getShapeKind(sh); getShapeId(sh); getShapeName(sh);
            getShapeText(sh); getShapePosition(sh); getShapeSize(sh);
            getShapeRotation(sh); getShapeFlip(sh);
          }
          if (ph) replaceTokensInSlide(s, { name: 'x' });
        }
        return await savePresentation(pres);
      }
    `);
    const dropped = [
      'addTable',
      'setTransition',
      'addLine',
      'addTextBox',
      'addShape',
      'setHyperlink',
      'setBullets',
      'setBackground',
      'addImage',
      'duplicateSlide',
    ];
    for (const name of dropped) {
      expect(fn.text, `bundle still references ${name}`).not.toContain(name);
    }
  });

  it('full fn-API bundle stays under 250 KB', async () => {
    // A consumer that uses every free-function entry point should ship
    // a bundle below 250 KB — wide upper bound to catch regressions.
    const fn = await bundleEntry(`
      import * as kit from 'pptx-kit';
      export async function run(bytes) {
        const pres = await kit.loadPresentation(bytes);
        const slides = kit.getSlides(pres);
        if (slides[0]) {
          const sh = kit.getSlideShapes(slides[0])[0];
          if (sh) {
            kit.setShapeText(sh, 'x');
            kit.setShapeFill(sh, '#FFFFFF');
            kit.setShapeStroke(sh, { color: '#000000' });
            kit.setShapePosition(sh, 1, 1);
            kit.setShapeSize(sh, 1, 1);
            kit.setShapeRotation(sh, 45);
          }
          kit.addSlideTextBox(slides[0], { x: 0, y: 0, w: 1, h: 1, text: 'x' });
          kit.addSlideShape(slides[0], { preset: 'rect', x: 0, y: 0, w: 1, h: 1 });
          kit.addSlideLine(slides[0], { from: { x: 0, y: 0 }, to: { x: 1, y: 1 } });
          kit.addSlideTable(slides[0], { x: 0, y: 0, w: 1, h: 1, rows: [['a']] });
          kit.addSlideImage(slides[0], new Uint8Array([0]), { x: 0, y: 0, w: 1, h: 1, format: 'png' });
          kit.setSlideBackground(slides[0], '#FF0000');
          kit.setSlideTransition(slides[0], { effect: 'fade' });
          kit.setSlideNotes(slides[0], 'x');
        }
        return await kit.savePresentation(pres);
      }
    `);
    expect(fn.bytes).toBeLessThan(250_000);
    process.stderr.write(`tree-shake: full-fn-API bundle = ${fn.bytes} bytes\n`);
  });

  it('deck-manipulation free-function bundle still drops authoring identifiers', async () => {
    const fn = await bundleEntry(`
      import {
        loadPresentation,
        savePresentation,
        getSlides,
        getSlideText,
        getSlideLayouts,
        addSlide,
        removeSlide,
        moveSlide,
        duplicateSlide,
        replaceTokensInPresentation,
      } from 'pptx-kit';
      export async function run(bytes) {
        const pres = await loadPresentation(bytes);
        const slides = getSlides(pres);
        if (slides[0]) getSlideText(slides[0]);
        for (const l of getSlideLayouts(pres)) {
          const s = addSlide(pres, { layout: l });
          if (slides[0]) {
            moveSlide(pres, s, 0);
            duplicateSlide(pres, s);
            removeSlide(pres, s);
          }
          break;
        }
        replaceTokensInPresentation(pres, { name: 'x' });
        return await savePresentation(pres);
      }
    `);
    const dropped = [
      'addTable',
      'setTransition',
      'addLine',
      'addTextBox',
      'addShape',
      'setHyperlink',
      'setBullets',
      'setBackground',
      'addImage',
    ];
    for (const name of dropped) {
      expect(fn.text, `bundle still references ${name}`).not.toContain(name);
    }
  });

  it('reports minimal-bundle size for the README record', async () => {
    const result = await bundleEntry(`
      import { loadPresentation, savePresentation } from 'pptx-kit';
      export async function run(bytes) {
        const pres = await loadPresentation(bytes);
        return await savePresentation(pres);
      }
    `);
    process.stderr.write(`tree-shake: minimal-bundle = ${result.bytes} bytes\n`);
    expect(result.bytes).toBeLessThan(120_000);
  });
});
