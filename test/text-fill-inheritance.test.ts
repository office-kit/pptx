import { describe, expect, it } from 'vitest';
import { resolveTextBodyRunFormatEffective } from '../src/api/fn/shape-paragraph.ts';
import { NS, parseXml } from '../src/internal/xml/index.ts';

const solid = '<a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>';
const pattern =
  '<a:pattFill prst="dkUpDiag"><a:fgClr><a:srgbClr val="112233"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>';

function effective(run: string, paragraph: string) {
  const body = parseXml(
    `<p:txBody xmlns:p="${NS.pml}" xmlns:a="${NS.dml}"><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr>${paragraph}</a:defRPr></a:pPr><a:r><a:rPr>${run}</a:rPr><a:t>Text</a:t></a:r></a:p></p:txBody>`,
  ).root;
  return resolveTextBodyRunFormatEffective({ theme: null }, body, 0, 0);
}

describe('text fill inheritance choices', () => {
  it('does not inherit a pattern over a directly applied solid font color', () => {
    const format = effective(solid, pattern);
    expect(format.color).toBe('#ABCDEF');
    expect(format.textFill).toBeUndefined();
  });

  it('does not inherit a solid color beside a directly applied pattern', () => {
    const format = effective(pattern, solid);
    expect(format.textFill?.kind).toBe('pattern');
    expect(format.color).toBeUndefined();
  });
});
