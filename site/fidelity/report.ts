// Writes the side-by-side HTML report: ground truth | our render | diff, with
// the SSIM / error numbers per slide. This is the human-facing companion to
// the machine-readable results.json — the same shape the playground dashboard
// will eventually grow into.

import { writeFileSync } from 'node:fs';
import { join } from 'node:path';

export interface SlideReport {
  readonly slide: number;
  readonly ssim: number;
  readonly fgSsim: number;
  readonly meanAbsError: number;
  readonly diffPercent: number;
  readonly gt: string | null;
  readonly ours: string;
  readonly diff: string | null;
}

export interface FileReport {
  readonly name: string;
  readonly meanSsim: number | null;
  readonly meanFgSsim: number | null;
  readonly slides: SlideReport[];
}

export interface ReportMeta {
  readonly engine: string;
  readonly width: number;
  readonly overallSsim: number | null;
  readonly overallFgSsim: number | null;
  readonly generatedNote: string;
}

const esc = (s: string): string =>
  s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);

const pct = (v: number): string => `${(v * 100).toFixed(1)}%`;
const ssimFmt = (v: number | null): string => (v === null ? 'n/a' : v.toFixed(4));

// Color a cell green→red by SSIM so a wall of slides is scannable.
const ssimColor = (v: number): string => {
  const clamped = Math.max(0, Math.min(1, (v - 0.5) / 0.5));
  const hue = Math.round(clamped * 120); // 0=red, 120=green
  return `hsl(${hue} 70% 88%)`;
};

const slideRow = (s: SlideReport): string => {
  const cell = (src: string | null): string =>
    src ? `<img loading="lazy" src="${esc(src)}" />` : '<div class="missing">—</div>';
  return `<tr>
    <td class="n">${s.slide}</td>
    <td class="metric" style="background:${ssimColor(s.fgSsim)}">
      <strong>${s.fgSsim.toFixed(4)}</strong> <span class="sub">fg</span><br/>
      <span class="sub">ssim ${s.ssim.toFixed(4)} · err ${pct(s.meanAbsError)} · diff ${pct(s.diffPercent)}</span>
    </td>
    <td>${cell(s.gt)}</td>
    <td>${cell(s.ours)}</td>
    <td>${cell(s.diff)}</td>
  </tr>`;
};

const fileSection = (f: FileReport): string => `
  <section>
    <h2>${esc(f.name)} <span class="mean">mean fg-SSIM ${ssimFmt(f.meanFgSsim)} · SSIM ${ssimFmt(f.meanSsim)}</span></h2>
    <table>
      <thead><tr><th>#</th><th>metrics</th><th>ground truth</th><th>pptx-kit</th><th>diff</th></tr></thead>
      <tbody>${f.slides.map(slideRow).join('')}</tbody>
    </table>
  </section>`;

export const writeReport = (outDir: string, meta: ReportMeta, files: FileReport[]): string => {
  const html = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"/>
<title>pptx-kit preview fidelity</title>
<style>
  :root { font-family: -apple-system, system-ui, sans-serif; }
  body { margin: 0; padding: 24px; background: #f7f7f8; color: #1f2330; }
  header { margin-bottom: 24px; }
  h1 { margin: 0 0 4px; font-size: 20px; }
  .meta { color: #6b7280; font-size: 13px; }
  .overall { font-size: 15px; margin-top: 8px; }
  section { background: #fff; border: 1px solid #e5e7eb; border-radius: 10px; padding: 16px; margin-bottom: 20px; }
  h2 { font-size: 15px; margin: 0 0 12px; display: flex; align-items: baseline; gap: 10px; }
  .mean { color: #6b7280; font-size: 13px; font-weight: 400; }
  table { border-collapse: collapse; width: 100%; }
  th { text-align: left; font-size: 12px; color: #6b7280; padding: 4px 8px; }
  td { padding: 6px 8px; vertical-align: top; border-top: 1px solid #f0f0f2; }
  td.n { font-weight: 600; color: #9ca3af; }
  td.metric { white-space: nowrap; font-size: 13px; border-radius: 6px; }
  td.metric .sub { color: #4b5563; font-size: 11px; }
  img { max-width: 320px; width: 100%; height: auto; border: 1px solid #e5e7eb; border-radius: 4px; background: #fff; display: block; }
  .missing { color: #d1d5db; font-size: 24px; text-align: center; }
</style></head>
<body>
  <header>
    <h1>pptx-kit preview fidelity</h1>
    <div class="meta">engine: ${esc(meta.engine)} · width: ${meta.width}px · ${esc(meta.generatedNote)}</div>
    <div class="overall">overall mean fg-SSIM: <strong>${ssimFmt(meta.overallFgSsim)}</strong> · plain SSIM: ${ssimFmt(meta.overallSsim)}</div>
  </header>
  ${files.map(fileSection).join('')}
</body></html>`;
  const path = join(outDir, 'index.html');
  writeFileSync(path, html);
  return path;
};
