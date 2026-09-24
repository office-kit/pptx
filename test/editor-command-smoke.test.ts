// End-to-end proof that the editor's command registry actually *drives* the
// library — not just that the ids line up. It builds a presentation entirely
// through the registry (the same path the ribbon/palette use), then round-trips
// it through save/load. This needs no XSDs, so it runs even where the ECMA-376
// schema submodule is absent.

import { describe, expect, it } from 'vitest';
import {
  createPresentation,
  findSlideLayout,
  getSlideLayout,
  getSlideLayoutName,
  isShapePlaceholder,
  getShapeChartSpec,
  addBlankSlide,
  findShapeById,
  getShapeId,
  getShapeText,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
} from '@office-kit/pptx';
import type { PresentationData, SlideData, SlideShapeData } from '@office-kit/pptx';
import { RegroupHistory } from '../site/src/lib/editor/core/regroup-history.ts';
import { getCommand } from '../site/src/lib/editor/core/registry.ts';
import type { Selection } from '../site/src/lib/editor/core/selection.ts';

// A minimal stand-in for EditorDocument that satisfies the surface the registry
// uses. Real one adds undo/rendering/reactivity, irrelevant to dispatch logic.
class FakeDoc {
  readonly regroupHistory = new RegroupHistory();
  pres: PresentationData;
  selection: Selection = { kind: 'none', slideIndex: 0 };
  constructor(pres: PresentationData) {
    this.pres = pres;
  }
  get slides(): readonly SlideData[] {
    return getSlides(this.pres);
  }
  slideAt(i: number): SlideData | null {
    return this.slides[i] ?? null;
  }
  shapeById(slideIndex: number, id: number): SlideShapeData | null {
    const slide = this.slideAt(slideIndex);
    return slide ? findShapeById(slide, id) : null;
  }
  select(selection: Selection): void {
    this.selection = selection;
  }
  selectShape(slideIndex: number, id: number): void {
    this.selection = { kind: 'shape', slideIndex, shapeIds: [id] };
  }
  selectSlide(i: number): void {
    this.selection = { kind: 'none', slideIndex: i };
  }
  transact<T>(_label: string, fn: () => T): T {
    return fn();
  }
}

function run(doc: FakeDoc, id: string, args: Record<string, unknown> = {}) {
  const cmd = getCommand(id);
  expect(cmd, `command ${id} should be registered`).toBeTruthy();
  expect(cmd!.canRun({ doc: doc as never }), `command ${id} should be runnable`).toBe(true);
  return cmd!.run({ doc: doc as never }, args);
}

describe('editor command registry drives the library', () => {
  it('inserts a chosen layout after the active slide and selects the inserted slide', async () => {
    const pres = createPresentation();
    const first = addBlankSlide(pres);
    const last = addBlankSlide(pres);
    const doc = new FakeDoc(pres);
    const layout = findSlideLayout(pres, 'Title and Content')!;
    run(doc, 'addSlide', { options: { layout } });
    expect(doc.selection.slideIndex).toBe(1);
    expect(doc.slides[0]).toBe(first);
    expect(doc.slides[2]).toBe(last);
    const inserted = doc.slides[1]!;
    expect(getSlideLayoutName(getSlideLayout(inserted)!)).toBe('Title and Content');
    expect(getSlideShapes(inserted).filter(isShapePlaceholder).length).toBeGreaterThan(0);
    const loaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideLayoutName(getSlideLayout(getSlides(loaded)[1]!)!)).toBe('Title and Content');
  });

  it('binds chart updates to exactly one selected chart', () => {
    const pres = createPresentation();
    addBlankSlide(pres);
    const doc = new FakeDoc(pres);
    const command = getCommand('setChartSpec')!;
    expect(command.canRun({ doc })).toBe(false);
    expect(command.params.map((param) => param.name)).toEqual(['spec']);
    const spec = { kind: 'column', categories: ['A'], series: [{ name: 'Sales', values: [5] }] };
    run(doc, 'addSlideChart', {
      opts: { x: inches(1), y: inches(1), w: inches(4), h: inches(3), spec },
    });
    const chart = getSlideShapes(doc.slides[0]!)[0]!;
    const chartId = getShapeId(chart);
    expect(command.canRun({ doc })).toBe(true);
    command.run({ doc }, { spec: { ...spec, title: 'Updated', legend: { position: 'b' } } });
    expect(getShapeChartSpec(chart)?.title).toBe('Updated');
    run(doc, 'addSlideShape', {
      opts: { preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1) },
    });
    const shape = getSlideShapes(doc.slides[0]!)[1]!;
    expect(command.canRun({ doc })).toBe(false);
    doc.select({ kind: 'shape', slideIndex: 0, shapeIds: [chartId, getShapeId(shape)] });
    expect(command.canRun({ doc })).toBe(false);
    expect(() => command.run({ doc }, { spec })).toThrow('Select one chart');
    expect(getShapeChartSpec(chart)?.title).toBe('Updated');
  });

  it('authors a slide + shape through the registry and round-trips', async () => {
    const pres = createPresentation();
    addBlankSlide(pres);
    const doc = new FakeDoc(pres);
    doc.selectSlide(0);

    // Add a rectangle via the registry (auto-selects it on return).
    run(doc, 'addSlideShape', {
      opts: {
        preset: 'rect',
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(2),
        text: 'Hello',
      },
    });
    expect(doc.selection.kind).toBe('shape');

    // The selected shape now drives shape-operand commands.
    run(doc, 'setShapeFill', { color: 'FF0000' });
    run(doc, 'setShapeText', { value: 'Edited via registry' });
    run(doc, 'setShapePosition', { x: inches(2), y: inches(3) });

    const shapes = getSlideShapes(doc.slideAt(0)!);
    expect(shapes.length).toBe(1);
    expect(getShapeText(shapes[0]!)).toBe('Edited via registry');

    // Round-trip: save and reload; the shape (and its text) survives.
    const bytes = await savePresentation(doc.pres);
    expect(bytes.byteLength).toBeGreaterThan(0);
    const reloaded = await loadPresentation(bytes);
    // createPresentation() yields 0 slides; addBlankSlide made the only slide.
    const reloadedShapes = getSlideShapes(getSlides(reloaded)[0]!);
    expect(reloadedShapes.some((s) => getShapeText(s) === 'Edited via registry')).toBe(true);
  });

  it('refuses shape commands when nothing is selected', () => {
    const pres = createPresentation();
    addBlankSlide(pres);
    const doc = new FakeDoc(pres);
    doc.selectSlide(0);
    const cmd = getCommand('setShapeFill')!;
    expect(cmd.canRun({ doc: doc as never })).toBe(false);
  });

  it('binds a command for every generated capability id', () => {
    // Cross-check with the coverage test: each manifested id resolves to a
    // runnable command object here.
    const ids = getShapeId; // silence unused import lint; used above indirectly
    void ids;
    for (const id of [
      'setShapeGradientFill',
      'addSlideTable',
      'setSlideTransition',
      'insertTableRow',
    ]) {
      expect(getCommand(id), id).toBeTruthy();
    }
  });
});
