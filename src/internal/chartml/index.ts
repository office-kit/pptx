// internal/chartml — c: namespace: charts + minimal xlsx writer for embedded data.
// Allowed imports: internal/drawingml, internal/parts, internal/xml.

export type { ChartKind, ChartSeries, ChartSpec } from './types.ts';
export { buildChartSpaceDoc } from './chart-builder.ts';
export type { DataRow } from './embedded-xlsx.ts';
export { buildEmbeddedXlsx, cellAddr, cellRange } from './embedded-xlsx.ts';
