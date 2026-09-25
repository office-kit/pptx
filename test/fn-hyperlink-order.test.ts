import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import {
  addSlideTextBox,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
  setShapeHyperlink,
  setShapeRunHyperlink,
  setShapeClickAction,
} from '../src/api/index.ts';
import { SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import {
  NS,
  elem,
  attr,
  parseFragment,
  qname,
  firstChildElement,
  type XmlElement,
} from '../src/internal/xml/index.ts';

const child = (parent: XmlElement, name: string, ns: string = NS.dml) =>
  firstChildElement(parent, qname('', name, ns))!;
const names = (parent: XmlElement) =>
  parent.children.flatMap((c) => (c.kind === 'element' ? [c.name.localName] : []));

it.each(['shape-url', 'range-url', 'run-url', 'shape-action', 'range-action'] as const)(
  '%s preserves schema order and imported hover, RTL and extension data through reload',
  async (mode) => {
    const pres = await loadPresentation(
      await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)),
    );
    const slide = getSlides(pres)[0]!;
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Link',
    });
    const properties = (root: XmlElement) =>
      mode === 'shape-action'
        ? child(child(root, 'nvSpPr', NS.pml), 'cNvPr', NS.pml)
        : child(child(child(root, 'txBody', NS.pml), 'p'), 'r');
    let parent = properties(shape[SHAPE_ELEMENT]);
    if (mode !== 'shape-action') {
      let rPr = firstChildElement(parent, qname('a', 'rPr', NS.dml));
      if (!rPr) {
        rPr = elem(qname('a', 'rPr', NS.dml));
        parent.children.unshift(rPr);
      }
      parent = rPr;
    }
    const suffix = [
      elem(qname('a', mode === 'shape-action' ? 'hlinkHover' : 'hlinkMouseOver', NS.dml), {
        attrs: [
          attr(qname('r', 'id', NS.officeDocRels), ''),
          attr(qname('', 'action', ''), 'ppaction://hlinkshowjump?jump=nextslide'),
        ],
      }),
    ];
    if (mode !== 'shape-action')
      suffix.push(elem(qname('a', 'rtl', NS.dml), { attrs: [attr(qname('', 'val', ''), '1')] }));
    suffix.push(
      elem(qname('a', 'extLst', NS.dml), {
        children: [
          elem(qname('a', 'ext', NS.dml), {
            attrs: [attr(qname('', 'uri', ''), 'urn:test:retained')],
            children: [parseFragment('<test:metadata xmlns:test="urn:test" value="保持"/>')],
          }),
        ],
      }),
    );
    parent.children.push(...structuredClone(suffix));
    const mutate = (clear = false) => {
      if (mode === 'shape-url') setShapeHyperlink(shape, clear ? null : 'https://example.com');
      else if (mode === 'range-url')
        setShapeHyperlink(shape, clear ? null : 'https://example.com', undefined, {
          range: { start: 0, end: 4 },
        });
      else if (mode === 'run-url')
        setShapeRunHyperlink(shape, 0, 0, clear ? null : 'https://example.com');
      else
        setShapeClickAction(
          shape,
          clear ? null : { kind: 'lastSlide' },
          mode === 'range-action' ? { range: { start: 0, end: 4 } } : undefined,
        );
    };
    for (const clear of [false, false, true]) {
      mutate(clear);
      const loaded = await loadPresentation(await savePresentation(pres));
      const result = getSlideShapes(getSlides(loaded)[0]!).at(-1)!;
      let resultParent = properties(result[SHAPE_ELEMENT]);
      if (mode !== 'shape-action') resultParent = child(resultParent, 'rPr');
      expect(names(resultParent)).toEqual([
        ...(clear ? [] : ['hlinkClick']),
        ...suffix.map((c) => c.name.localName),
      ]);
      expect(
        resultParent.children.filter(
          (c) => c.kind === 'element' && c.name.localName !== 'hlinkClick',
        ),
      ).toEqual(suffix);
    }
  },
);
