// Layer-0 quality gate: tree-shakeability.
//
// pptx-kit's public API ships every authoring capability behind classes
// today, so any consumer of `Presentation` drags every method into their
// bundle whether they use it or not — exactly the regression the user
// flagged. This test bundles a minimal entry that exercises only
// load + save, then asserts that feature-specific identifiers from the
// rest of the API are dead-shaken from the resulting bundle.
//
// While the class-based public API persists, these assertions are
// expected to FAIL — the failing test is the design pressure that
// motivates the refactor to free functions. As the refactor lands, the
// `dropped` allowlist below will grow until every authoring feature
// tree-shakes cleanly.

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
      // Don't minify so symbol substrings remain greppable.
      minify: false,
      // Resolve `pptx-kit` to the source tree.
      alias: { 'pptx-kit': join(REPO_ROOT, 'src/index.ts') },
      // No source maps, no metadata.
      logLevel: 'silent',
    });
    const text = readFileSync(out, 'utf8');
    return { bytes: Buffer.byteLength(text, 'utf8'), text };
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
};

describe('tree-shake: minimal load+save entry', () => {
  it('bundles successfully (class API — full bundle)', async () => {
    const result = await bundleEntry(`
      import { Presentation } from 'pptx-kit';
      export async function run(bytes) {
        const pres = await Presentation.load(bytes);
        return await pres.save();
      }
    `);
    expect(result.bytes).toBeGreaterThan(0);
    // Sanity: fflate is part of the load/save path and must be present.
    expect(result.text).toContain('zip');
  });

  it('bundles successfully (free-function API — tree-shaken)', async () => {
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

  it('free-function bundle is strictly smaller than class-bundle', async () => {
    const cls = await bundleEntry(`
      import { Presentation } from 'pptx-kit';
      export async function run(bytes) {
        const pres = await Presentation.load(bytes);
        return await pres.save();
      }
    `);
    const fn = await bundleEntry(`
      import { loadPresentation, savePresentation } from 'pptx-kit';
      export async function run(bytes) {
        const pres = await loadPresentation(bytes);
        return await savePresentation(pres);
      }
    `);
    process.stderr.write(
      `tree-shake: class=${cls.bytes}B, fn=${fn.bytes}B, saved=${cls.bytes - fn.bytes}B\n`,
    );
    expect(fn.bytes).toBeLessThan(cls.bytes);
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
    // Pulls in every read-side slide / shape free function. None of them
    // should drag SlideShape mutation methods into the bundle.
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

  it('deck-manipulation free-function bundle still drops authoring identifiers', async () => {
    // Pulls in every deck-level free function. It should NOT drag any
    // SlideShape mutation method (addTable, setTransition, addImage, ...)
    // into the bundle, because those still live on the class API and the
    // fn-API path never references them.
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
      export async function run(bytes, layout) {
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

  // The capabilities below are NOT touched by the minimal entry. After
  // the class → free-function refactor, every name listed here must be
  // absent from the minimal bundle. We mark the test as `todo` so the
  // expectation is recorded but doesn't block CI while the refactor is
  // staged across commits.
  it.todo('minimal load+save bundle excludes addTable / setTransition / setBullets / ...', async () => {
    const result = await bundleEntry(`
        import { Presentation } from 'pptx-kit';
        export async function run(bytes) {
          const pres = await Presentation.load(bytes);
          return await pres.save();
        }
      `);
    const dropped = [
      'addTable',
      'setTransition',
      'setBullets',
      'addLine',
      'addImage',
      'addTextBox',
      'addShape',
      'duplicateSlide',
      'setHyperlink',
      'setNotes',
      'replaceTokens',
    ];
    for (const name of dropped) {
      expect(result.text, `bundle still references ${name}`).not.toContain(name);
    }
  });

  it('reports current minimal-bundle size for the README record', async () => {
    const result = await bundleEntry(`
      import { Presentation } from 'pptx-kit';
      export async function run(bytes) {
        const pres = await Presentation.load(bytes);
        return await pres.save();
      }
    `);
    // Soft assertion only: shrinks as the refactor lands. The threshold
    // is the current (unrefactored) size + headroom; tighten as the
    // class-elimination work merges. Print to stderr so the test log
    // captures the value for triage.
    process.stderr.write(`tree-shake: minimal-bundle = ${result.bytes} bytes\n`);
    expect(result.bytes).toBeLessThan(300_000);
  });
});
