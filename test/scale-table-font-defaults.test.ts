import { expect, it } from 'vitest';
import { scaleSlideContent } from '../src/api/fn/_scale-slide-content.ts';
import { parseXml, serializeXml } from '../src/internal/xml/index.ts';

it('preserves table font inheritance and explicit run, field, break and end sizes', () => {
  const doc = parseXml(
    `<a:tc xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:txBody><a:bodyPr/><a:lstStyle><a:lvl2pPr><a:defRPr sz="2200"/></a:lvl2pPr></a:lstStyle><a:p><a:pPr lvl="1"/><a:r><a:t>Level</a:t></a:r></a:p><a:p><a:pPr><a:defRPr sz="2400"/></a:pPr><a:r><a:t>Paragraph</a:t></a:r></a:p><a:p><a:r><a:rPr sz="3000"/><a:t>Explicit</a:t></a:r><a:br><a:rPr sz="1200"/></a:br><a:fld id="{11111111-1111-1111-1111-111111111111}" type="slidenum"><a:t>1</a:t></a:fld><a:endParaRPr sz="2000"/></a:p></a:txBody></a:tc>`,
  );
  scaleSlideContent(doc.root, 2, 0, 0);
  const xml = serializeXml(doc);
  expect(xml).toContain('<a:defRPr sz="4400"');
  expect(xml).toContain('<a:defRPr sz="4800"');
  expect(xml).toContain('<a:rPr sz="6000"');
  expect(xml).toContain('<a:rPr sz="2400"');
  expect(xml).toContain('<a:endParaRPr sz="4000"');
  expect(xml).not.toContain('sz="3600"');
  expect(xml).toContain('<a:r><a:t>Level</a:t></a:r>');
  expect(xml).toContain('<a:r><a:t>Paragraph</a:t></a:r>');
});
