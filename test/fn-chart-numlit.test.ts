// Regression: chart values authored as `<c:numLit>` (inline literal
// array) instead of `<c:numRef>` (reference into the embedded
// workbook). @office-kit/pptx's own writer always emits `numRef` with a cache,
// but charts produced by python-pptx, hand-edited XML, or older
// pptxgenjs paths sometimes use `numLit`. Before this fix, the reader
// only walked `numRef`, so the values came back as empty arrays and
// any renderer (the playground SVG one included) drew nothing.

import { describe, expect, it } from 'vitest';
import { readChartSpec } from '../src/internal/chartml/index.ts';
import { parseXml } from '../src/internal/xml/index.ts';

const CHART_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx>
            <c:strLit>
              <c:ptCount val="1"/>
              <c:pt idx="0"><c:v>Revenue</c:v></c:pt>
            </c:strLit>
          </c:tx>
          <c:cat>
            <c:strLit>
              <c:ptCount val="4"/>
              <c:pt idx="0"><c:v>Q1</c:v></c:pt>
              <c:pt idx="1"><c:v>Q2</c:v></c:pt>
              <c:pt idx="2"><c:v>Q3</c:v></c:pt>
              <c:pt idx="3"><c:v>Q4</c:v></c:pt>
            </c:strLit>
          </c:cat>
          <c:val>
            <c:numLit>
              <c:formatCode>General</c:formatCode>
              <c:ptCount val="4"/>
              <c:pt idx="0"><c:v>120</c:v></c:pt>
              <c:pt idx="1"><c:v>180</c:v></c:pt>
              <c:pt idx="2"><c:v>240</c:v></c:pt>
              <c:pt idx="3"><c:v>300</c:v></c:pt>
            </c:numLit>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

describe('chart reader: literal value channels', () => {
  it('reads `<c:strLit>` / `<c:numLit>` like their `Ref` siblings', () => {
    const root = parseXml(CHART_XML).root;
    const spec = readChartSpec(root);
    expect(spec).not.toBeNull();
    if (!spec) return;
    expect(spec.kind).toBe('column');
    expect(spec.categories).toEqual(['Q1', 'Q2', 'Q3', 'Q4']);
    expect(spec.series).toHaveLength(1);
    expect(spec.series[0]?.name).toBe('Revenue');
    expect(spec.series[0]?.values).toEqual([120, 180, 240, 300]);
  });
});
