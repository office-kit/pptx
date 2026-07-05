// Ground-truth rasterization: turn a .pptx into one reference image per slide
// using a real presentation engine, so our renderer can be measured against
// "what a user actually sees".
//
// Tiered, per the preview-fidelity roadmap:
//   - libreoffice (default): headless `soffice`, free, CI-friendly. The gate.
//   - powerpoint (opt-in, local, macOS): true ground truth via AppleScript.
//
// Both routes go .pptx -> .pdf -> .ppm: PDF is the one export format both
// engines produce reliably from the CLI, and `pdftoppm` (poppler) gives us a
// dependency-light raster. Select with the GROUND_TRUTH env var.

import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join } from 'node:path';
import { decodePpm } from './ppm.ts';
import type { RgbaImage } from './image.ts';

export type GroundTruthEngine = 'libreoffice' | 'powerpoint';

export interface GroundTruthOptions {
  readonly width: number;
  readonly height: number;
  readonly engine: GroundTruthEngine;
}

export class GroundTruthUnavailableError extends Error {}

// Candidate locations for the LibreOffice binary, in priority order. The env
// override lets CI point at an apt-installed `soffice` directly.
const sofficeCandidates = (): string[] =>
  [
    process.env.PPTX_KIT_SOFFICE,
    'soffice',
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    '/usr/bin/soffice',
    '/usr/bin/libreoffice',
  ].filter((c): c is string => Boolean(c));

export const resolveSoffice = (): string => {
  for (const c of sofficeCandidates()) {
    try {
      // `soffice --version` is the cheapest liveness probe.
      execFileSync(c, ['--version'], { stdio: 'ignore' });
      return c;
    } catch {
      // try next
    }
  }
  throw new GroundTruthUnavailableError(
    'LibreOffice not found. Install it (macOS: `brew install --cask libreoffice`; ' +
      'Debian/CI: `apt-get install libreoffice`) or set PPTX_KIT_SOFFICE.',
  );
};

const resolvePdftoppm = (): string => {
  const candidate = process.env.PPTX_KIT_PDFTOPPM ?? 'pdftoppm';
  try {
    execFileSync(candidate, ['-v'], { stdio: 'ignore' });
  } catch {
    throw new GroundTruthUnavailableError(
      'pdftoppm (poppler) not found. Install it (macOS: `brew install poppler`; ' +
        'Debian/CI: `apt-get install poppler-utils`) or set PPTX_KIT_PDFTOPPM.',
    );
  }
  return candidate;
};

// A LibreOffice headless run needs a private user profile, or a second
// invocation collides with "soffice is already running". One persistent
// profile per machine is enough since the harness runs sequentially.
const loProfileDir = (): string => {
  const dir = join(tmpdir(), 'office-kit-pptx-fidelity-lo-profile');
  mkdirSync(dir, { recursive: true });
  return dir;
};

const pptxToPdfLibreOffice = (pptxPath: string, outDir: string): string => {
  const soffice = resolveSoffice();
  execFileSync(
    soffice,
    [
      '--headless',
      '--norestore',
      '--nologo',
      `-env:UserInstallation=file://${loProfileDir()}`,
      '--convert-to',
      'pdf:impress_pdf_Export',
      '--outdir',
      outDir,
      pptxPath,
    ],
    { stdio: 'ignore', timeout: 120_000 },
  );
  const pdf = join(outDir, basename(pptxPath).replace(/\.pptx$/i, '.pdf'));
  if (!existsSync(pdf)) throw new Error(`LibreOffice produced no PDF for ${pptxPath}`);
  return pdf;
};

// macOS-only, opt-in. PowerPoint exports the deck to PDF via AppleScript; the
// rest of the pipeline is shared. Untested in CI (no PowerPoint there) — it is
// the local high-fidelity check the roadmap calls for.
const pptxToPdfPowerPoint = (pptxPath: string, outDir: string): string => {
  if (process.platform !== 'darwin') {
    throw new GroundTruthUnavailableError('PowerPoint ground truth is macOS-only.');
  }
  const pdf = join(outDir, basename(pptxPath).replace(/\.pptx$/i, '.pdf'));
  // PowerPoint's `open`/`save in` expect a file reference, not a raw POSIX
  // path string — passing the bare string silently fails to produce a PDF.
  // `POSIX file (...)` coerces the argv string into the file object PP wants.
  const script = [
    'on run argv',
    '  set inPath to item 1 of argv',
    '  set outPath to item 2 of argv',
    '  tell application "Microsoft PowerPoint"',
    // `activate` foregrounds PowerPoint before the open/save. Without it,
    // a backgrounded PowerPoint intermittently hangs the AppleScript
    // (osascript ETIMEDOUT) instead of exporting.
    '    activate',
    '    open (POSIX file inPath)',
    '    set theDoc to active presentation',
    '    save theDoc in (POSIX file outPath) as save as PDF',
    '    close theDoc saving no',
    '  end tell',
    'end run',
  ].join('\n');
  try {
    execFileSync('osascript', ['-e', script, pptxPath, pdf], {
      stdio: 'ignore',
      timeout: 120_000,
    });
  } catch (err) {
    throw new GroundTruthUnavailableError(
      `PowerPoint export failed (is PowerPoint installed and scriptable?): ${String(err)}`,
    );
  }
  if (!existsSync(pdf)) throw new Error(`PowerPoint produced no PDF for ${pptxPath}`);
  return pdf;
};

const pdfToImages = (
  pdfPath: string,
  outDir: string,
  width: number,
  height: number,
): RgbaImage[] => {
  const pdftoppm = resolvePdftoppm();
  const prefix = join(outDir, 'page');
  // -scale-to-x / -scale-to-y together force an exact pixel box, so every
  // slide image already matches the dimensions our renderer targets.
  execFileSync(
    pdftoppm,
    ['-scale-to-x', String(width), '-scale-to-y', String(height), pdfPath, prefix],
    { stdio: 'ignore', timeout: 120_000 },
  );
  // pdftoppm zero-pads the page index to the page-count width; sort numerically.
  const files = readdirSync(outDir)
    .filter((f) => f.startsWith('page') && f.endsWith('.ppm'))
    .sort((a, b) => pageNum(a) - pageNum(b));
  return files.map((f) => decodePpm(readFileSync(join(outDir, f))));
};

const pageNum = (file: string): number => {
  const m = file.match(/page-?(\d+)\.ppm$/);
  return m ? Number(m[1]) : 0;
};

/** Render every slide of `pptxPath` to a reference image. */
export const renderGroundTruth = (pptxPath: string, opts: GroundTruthOptions): RgbaImage[] => {
  const work = mkdtempSync(join(tmpdir(), 'office-kit-pptx-gt-'));
  try {
    const pdf =
      opts.engine === 'powerpoint'
        ? pptxToPdfPowerPoint(pptxPath, work)
        : pptxToPdfLibreOffice(pptxPath, work);
    return pdfToImages(pdf, work, opts.width, opts.height);
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
};
