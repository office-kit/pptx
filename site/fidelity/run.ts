// Preview-fidelity harness entrypoint.
//
//   pnpm --filter pptx-kit-site fidelity                 # all samples, LibreOffice
//   pnpm --filter pptx-kit-site fidelity -- --ours-only  # skip ground truth
//   GROUND_TRUTH=powerpoint pnpm --filter ... fidelity   # local PP check (macOS)
//
// Produces site/fidelity/out/{index.html, results.json} plus per-slide PNGs:
// ground truth, our render, and the diff. See site/fidelity/README.md.

import { mkdirSync, readdirSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { resizeRgba } from './image.ts';
import { encodePng } from './png.ts';
import { compareImages, diffImage } from './ssim.ts';
import { renderPresentation } from './render-ours.ts';
import {
  GroundTruthUnavailableError,
  renderGroundTruth,
  type GroundTruthEngine,
} from './ground-truth.ts';
import { writeReport, type FileReport, type SlideReport } from './report.ts';
import {
  compareWithBaseline,
  getSofficeVersion,
  loadBaseline,
  reportsToRunFiles,
  writeBaselineFile,
  type BaselineData,
} from './baseline.ts';

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(HERE, '..', '..');
const BASELINE_PATH = join(HERE, 'baseline.json');
const CANDIDATE_PATH = join(HERE, 'baseline.candidate.json');

interface Args {
  readonly width: number;
  readonly engine: GroundTruthEngine;
  readonly outDir: string;
  readonly samplesDir: string;
  readonly oursOnly: boolean;
  readonly check: boolean;
  readonly record: boolean;
  readonly files: string[];
}

const parseArgs = (argv: string[]): Args => {
  let width = 1280;
  let engine: GroundTruthEngine =
    process.env.GROUND_TRUTH === 'powerpoint' ? 'powerpoint' : 'libreoffice';
  let outDir = join(HERE, 'out');
  let samplesDir = join(REPO_ROOT, 'samples', 'out');
  let oursOnly = false;
  let check = false;
  let record = false;
  const files: string[] = [];
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i]!;
    if (a === '--width') width = Number(argv[++i]);
    else if (a === '--engine') engine = argv[++i] === 'powerpoint' ? 'powerpoint' : 'libreoffice';
    else if (a === '--out') outDir = resolve(argv[++i]!);
    else if (a === '--samples') samplesDir = resolve(argv[++i]!);
    else if (a === '--ours-only') oursOnly = true;
    else if (a === '--check') check = true;
    else if (a === '--record') record = true;
    else if (!a.startsWith('--')) files.push(resolve(a));
  }
  return { width, engine, outDir, samplesDir, oursOnly, check, record, files };
};

const discoverSamples = (dir: string): string[] => {
  let entries: string[];
  try {
    entries = readdirSync(dir);
  } catch {
    return [];
  }
  return entries
    .filter((f) => f.toLowerCase().endsWith('.pptx'))
    .sort()
    .map((f) => join(dir, f));
};

const mean = (xs: number[]): number | null =>
  xs.length === 0 ? null : xs.reduce((a, b) => a + b, 0) / xs.length;

const processFile = (
  pptxPath: string,
  args: Args,
  oursRender: Awaited<ReturnType<typeof renderPresentation>>,
): FileReport => {
  const name = pptxPath.split('/').pop() ?? pptxPath;
  const fileOut = join(args.outDir, name.replace(/\.pptx$/i, ''));
  mkdirSync(fileOut, { recursive: true });

  let gtImages: ReturnType<typeof renderGroundTruth> | null = null;
  if (!args.oursOnly) {
    gtImages = renderGroundTruth(pptxPath, {
      width: oursRender.pixelWidth,
      height: oursRender.pixelHeight,
      engine: args.engine,
    });
    if (gtImages.length !== oursRender.slides.length) {
      console.warn(
        `  ! slide-count mismatch for ${name}: ground truth ${gtImages.length} vs ours ${oursRender.slides.length} (hidden slides?). Comparing the overlap.`,
      );
    }
  }

  const count = gtImages
    ? Math.min(gtImages.length, oursRender.slides.length)
    : oursRender.slides.length;
  const slides: SlideReport[] = [];
  for (let i = 0; i < count; i++) {
    const ours = oursRender.slides[i]!;
    const oursRel = `${name.replace(/\.pptx$/i, '')}/slide-${i + 1}.ours.png`;
    writeFileSync(join(args.outDir, oursRel), ours.png);

    if (!gtImages) {
      slides.push({
        slide: i + 1,
        ssim: 1,
        fgSsim: 1,
        meanAbsError: 0,
        diffPercent: 0,
        gt: null,
        ours: oursRel,
        diff: null,
      });
      continue;
    }

    const gt = gtImages[i]!;
    const oursMatched = resizeRgba(ours.image, gt.width, gt.height);
    const result = compareImages(gt, oursMatched);
    const diff = diffImage(gt, oursMatched);
    const gtRel = `${name.replace(/\.pptx$/i, '')}/slide-${i + 1}.gt.png`;
    const diffRel = `${name.replace(/\.pptx$/i, '')}/slide-${i + 1}.diff.png`;
    writeFileSync(join(args.outDir, gtRel), encodePng(gt));
    writeFileSync(join(args.outDir, diffRel), encodePng(diff));
    slides.push({
      slide: i + 1,
      ssim: result.ssim,
      fgSsim: result.fgSsim,
      meanAbsError: result.meanAbsError,
      diffPercent: result.diffPercent,
      gt: gtRel,
      ours: oursRel,
      diff: diffRel,
    });
  }

  const meanSsim = args.oursOnly ? null : mean(slides.map((s) => s.ssim));
  const meanFgSsim = args.oursOnly ? null : mean(slides.map((s) => s.fgSsim));
  return { name, meanSsim, meanFgSsim, slides };
};

const main = async (): Promise<void> => {
  const args = parseArgs(process.argv.slice(2));

  if ((args.check || args.record) && args.oursOnly) {
    console.error('--check and --record cannot be combined with --ours-only.');
    process.exit(1);
  }
  if (args.check && args.record) {
    console.error('--check and --record cannot be combined with each other.');
    process.exit(1);
  }

  const pptxFiles = args.files.length > 0 ? args.files : discoverSamples(args.samplesDir);
  if (pptxFiles.length === 0) {
    console.error(
      `No .pptx files found (looked in ${args.samplesDir}). Run \`pnpm samples\` first.`,
    );
    process.exit(2);
  }
  mkdirSync(args.outDir, { recursive: true });
  console.log(
    `Fidelity run: ${pptxFiles.length} file(s), width ${args.width}px, ` +
      `${args.oursOnly ? 'ours-only (no ground truth)' : `engine ${args.engine}`}`,
  );

  const reports: FileReport[] = [];
  for (const pptx of pptxFiles) {
    const shortName = pptx.split('/').pop();
    process.stdout.write(`- ${shortName} … `);
    try {
      const oursRender = await renderPresentation(pptx, { width: args.width });
      const report = processFile(pptx, args, oursRender);
      reports.push(report);
      const m = report.meanFgSsim;
      console.log(
        `${report.slides.length} slide(s)${m === null ? '' : `, fg-SSIM ${m.toFixed(4)}`}`,
      );
    } catch (err) {
      if (err instanceof GroundTruthUnavailableError) {
        console.log('\n');
        console.error(err.message);
        console.error('Re-run with `--ours-only` to skip ground truth.');
        process.exit(2);
      }
      throw err;
    }
  }

  const allSlides = reports.flatMap((r) => r.slides);
  const overall = args.oursOnly ? null : mean(allSlides.map((s) => s.ssim));
  const overallFg = args.oursOnly ? null : mean(allSlides.map((s) => s.fgSsim));
  const note = `${reports.length} file(s), ${reports.reduce((n, r) => n + r.slides.length, 0)} slide(s)`;
  writeReport(
    args.outDir,
    {
      engine: args.engine,
      width: args.width,
      overallSsim: overall,
      overallFgSsim: overallFg,
      generatedNote: note,
    },
    reports,
  );
  writeFileSync(
    join(args.outDir, 'results.json'),
    JSON.stringify(
      {
        engine: args.engine,
        width: args.width,
        overallSsim: overall,
        overallFgSsim: overallFg,
        files: reports,
      },
      null,
      2,
    ),
  );

  console.log('');
  console.log(
    `overall mean fg-SSIM: ${overallFg === null ? 'n/a (ours-only)' : overallFg.toFixed(4)}` +
      `${overall === null ? '' : ` · plain SSIM: ${overall.toFixed(4)}`}`,
  );
  console.log(`report: ${join(args.outDir, 'index.html')}`);

  if (args.check || args.record) {
    const runFiles = reportsToRunFiles(reports);
    const sofficeVersion = args.engine === 'libreoffice' ? getSofficeVersion() : 'n/a';
    const candidateData: BaselineData = {
      engine: args.engine,
      width: args.width,
      sofficeVersion,
      files: runFiles,
    };

    if (args.record) {
      writeBaselineFile(BASELINE_PATH, candidateData);
      console.log(`\nBaseline recorded → ${BASELINE_PATH}`);
      return;
    }

    // --check: always write the candidate so CI can upload it as an artifact
    // regardless of pass/fail, then gate against the committed baseline.
    writeBaselineFile(CANDIDATE_PATH, candidateData);
    console.log(`\nBaseline candidate → ${CANDIDATE_PATH}`);

    const baselineData = loadBaseline(BASELINE_PATH);
    if (baselineData === null) {
      console.error(`\nNo baseline found at ${BASELINE_PATH}.`);
      console.error('Run with --record to create one locally, or in CI download the');
      console.error(
        'fidelity-baseline-candidate artifact and commit it as site/fidelity/baseline.json.',
      );
      process.exit(1);
    }

    const cmp = compareWithBaseline(baselineData, { files: runFiles });

    if (cmp.improvements.length > 0) {
      console.log('\nImprovements (fg-SSIM rose > 0.02 above baseline):');
      for (const imp of cmp.improvements) {
        console.log(
          `  ${imp.file} slide ${imp.slide}: ${imp.baseline.toFixed(4)} → ${imp.actual.toFixed(4)} (+${imp.delta.toFixed(4)})`,
        );
      }
      console.log('  Tip: re-record with --record so the baseline tracks progress.');
    }

    if (cmp.pass) {
      console.log('\nFidelity gate PASSED.');
    } else {
      console.error('\nFidelity gate FAILED:');
      console.error(
        `  ${'FILE'.padEnd(38)} ${'SLIDE'.padEnd(6)} ${'REASON'.padEnd(22)} ${'ACTUAL'.padEnd(8)} ${'BASELINE'.padEnd(9)} DELTA`,
      );
      for (const f of cmp.failures) {
        const slideStr = f.slide === 0 ? '—' : String(f.slide);
        const actualStr = f.actual !== undefined ? f.actual.toFixed(4) : '—';
        const baseStr = f.baseline !== undefined ? f.baseline.toFixed(4) : '—';
        const deltaStr = f.delta !== undefined ? f.delta.toFixed(4) : '—';
        console.error(
          `  ${f.file.padEnd(38)} ${slideStr.padEnd(6)} ${f.reason.padEnd(22)} ${actualStr.padEnd(8)} ${baseStr.padEnd(9)} ${deltaStr}`,
        );
      }
      process.exit(1);
    }
  }
};

main().catch((err: unknown) => {
  console.error(err);
  process.exit(1);
});
