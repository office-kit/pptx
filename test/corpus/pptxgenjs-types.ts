// Minimal structural types for the slice of the PptxGenJS API the corpus
// exercises. PptxGenJS ships no types we can import from the submodule's CJS
// bundle, and the project lints `no-explicit-any`, so we hand-declare just
// enough surface to author cases without `any`.

export type PgjsOpts = Record<string, unknown>;

export interface PgjsChartSeries {
  name: string;
  labels: string[];
  values: number[];
}

export interface PgjsSlide {
  addText(text: string, opts: PgjsOpts): void;
  addShape(type: string, opts: PgjsOpts): void;
  addTable(rows: string[][], opts: PgjsOpts): void;
  addImage(opts: PgjsOpts): void;
  addChart(type: string, data: PgjsChartSeries[], opts: PgjsOpts): void;
}

export interface PgjsDeck {
  addSlide(): PgjsSlide;
  write(outputType: 'nodebuffer'): Promise<Uint8Array>;
}

export type PgjsCtor = new () => PgjsDeck;
