import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as pptx from '@office-kit/pptx';
import {
  editorModel,
  recordEdit,
  replayEdits,
  emptyHistory,
  restoreEditorSession,
} from '../src/editor.ts';
import { shapeCategories, lineTools } from '../src/shape-gallery.ts';
import {
  getPresetShapePath,
  renderSlideToSvg,
  measureTableCellHeight,
} from '@office-kit/pptx-preview';
import { replaceParagraphRange, resizedTextBounds } from '../src/rich-text.ts';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
async function fixture() {
  const p = pptx.createPresentation();
  const slide = pptx.addBlankSlide(p);
  pptx.addSlideTextBox(slide, {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(1),
    text: 'Original',
  });
  return pptx.loadPresentation(await pptx.savePresentation(p));
}
test('kept show ink survives export and replays as one multi-slide undo step', async () => {
  const p = await fixture();
  pptx.addBlankSlide(p);
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const stroke = {
    points: [
      { x: 914400, y: 914400 },
      { x: 1828800, y: 1828800 },
    ],
    widthEmu: 25400,
    color: '#FF0000',
  };
  const command = {
    type: 'ink',
    slide: 0,
    ink: before.slides.map((slide) => ({ slide: slide.key, strokes: [stroke] })),
  };
  const history = emptyHistory();
  history.entries.push(recordEdit(p, command));
  history.cursor = 1;
  const exported = await pptx.loadPresentation(await pptx.savePresentation(p));
  for (const slide of pptx.getSlides(exported)) {
    const ink = pptx.getSlideShapes(slide).at(-1);
    assert.equal(pptx.getShapeKind(ink), 'ink');
    assert.ok(pptx.getShapeImageBytes(ink).length > 100);
  }
  const redo = await pptx.loadPresentation(base);
  replayEdits(redo, history);
  assert.deepEqual(editorModel(redo), editorModel(exported));
  history.cursor = 0;
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, history);
  assert.deepEqual(editorModel(undo), before);
  const invalid = await pptx.loadPresentation(base);
  assert.throws(() =>
    recordEdit(invalid, {
      ...command,
      ink: [command.ink[0], { ...command.ink[1], strokes: [{ ...stroke, color: 'invalid' }] }],
    }),
  );
  assert.deepEqual(editorModel(invalid), before);
  history.cursor = 1;
  pptx.addSlideTextBox(pptx.getSlides(undo)[1], {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(1),
    h: pptx.inches(1),
    text: 'Changed target',
  });
  assert.throws(() => replayEdits(undo, history), /conflict|changed/i);
});
test('editing frame preserves authored margins, anchoring, wrapping and columns', async () => {
  const p = await fixture();
  const shape = pptx.getSlideShapes(pptx.getSlides(p)[0])[0];
  pptx.setShapeTextMargins(shape, { left: 0, top: 12000, right: 34000, bottom: 56000 });
  pptx.setShapeTextAnchor(shape, 'bottom');
  pptx.setShapeTextWrap(shape, 'none');
  pptx.setShapeTextColumns(shape, { count: 2, gapEmu: 78000 });
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(saved).slides[0].shapes[0].textFrame, {
    margins: { left: 0, top: 12000, right: 34000, bottom: 56000 },
    anchor: 'bottom',
    wrap: false,
    columns: { count: 2, gapEmu: 78000 },
  });
});
test('text box properties persist, replay, undo and detect changed source margins', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const id = before.slides[0].shapes[0].id;
  const h = emptyHistory();
  h.entries.push(
    recordEdit(p, {
      type: 'update',
      slide: 0,
      ids: [id],
      changes: {
        margins: { left: 180000, top: 36000, right: 72000, bottom: 108000 },
        wrap: false,
        anchor: 'bottom',
      },
    }),
  );
  h.cursor = 1;
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(saved).slides[0].shapes[0].textFrame, {
    margins: { left: 180000, top: 36000, right: 72000, bottom: 108000 },
    wrap: false,
    anchor: 'bottom',
    columns: null,
  });
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, h);
  assert.deepEqual(editorModel(replayed), editorModel(saved));
  const conflict = await pptx.loadPresentation(base);
  pptx.setShapeTextMargins(pptx.getSlideShapes(pptx.getSlides(conflict)[0])[0], { left: 1 });
  assert.throws(() => replayEdits(conflict, h), /conflict/i);
  h.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, h);
  assert.deepEqual(editorModel(undone), before);
  for (const changes of [
    { margins: { left: -1 } },
    { margins: { right: 20116801 } },
    { wrap: 'yes' },
  ]) {
    assert.throws(() => recordEdit(undone, { type: 'update', slide: 0, ids: [id], changes }));
  }
});
test('columns dialog edits preserve mixed values and survive export, replay and undo', async () => {
  const p = await fixture();
  const slide = pptx.getSlides(p)[0];
  const first = pptx.getSlideShapes(slide)[0];
  const second = pptx.addSlideTextBox(slide, {
    x: 0,
    y: 0,
    w: 3000000,
    h: 1000000,
    text: 'Second',
  });
  pptx.setShapeTextColumns(first, { count: 2, gapEmu: 180000 });
  pptx.setShapeTextColumns(second, { count: 3, gapEmu: 360000 });
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const ids = before.slides[0].shapes.map((shape) => shape.id);
  const h = emptyHistory();
  h.entries.push(
    recordEdit(p, { type: 'update', slide: 0, ids, changes: { textColumns: { count: 1 } } }),
  );
  h.cursor++;
  assert.deepEqual(
    editorModel(p).slides[0].shapes.map((shape) => shape.textFrame.columns),
    [
      { count: 1, gapEmu: 180000 },
      { count: 1, gapEmu: 360000 },
    ],
  );
  h.entries.push(
    recordEdit(p, { type: 'update', slide: 0, ids, changes: { textColumns: { gapEmu: 72000 } } }),
  );
  h.cursor++;
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(
    editorModel(saved).slides[0].shapes.map((shape) => shape.textFrame.columns),
    [
      { count: 1, gapEmu: 72000 },
      { count: 1, gapEmu: 72000 },
    ],
  );
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, h);
  assert.deepEqual(editorModel(replayed), editorModel(saved));
  const conflict = await pptx.loadPresentation(base);
  pptx.setShapeTextColumns(pptx.getSlideShapes(pptx.getSlides(conflict)[0])[0], {
    count: 2,
    gapEmu: 1,
  });
  assert.throws(() => replayEdits(conflict, h), /conflict/i);
  h.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, h);
  assert.deepEqual(editorModel(undone), before);
  for (const textColumns of [
    { count: 0 },
    { count: 17 },
    { count: 1.5 },
    { gapEmu: -1 },
    { gapEmu: 14630401 },
  ])
    assert.throws(() =>
      recordEdit(undone, { type: 'update', slide: 0, ids, changes: { textColumns } }),
    );
});
test('editor exposes inherited autofit scales and respects explicit mode overrides', async () => {
  const original = pptx.createPresentation();
  pptx.addTitleSlide(original, 'Inherited text');
  const parts = unzipSync(await pptx.savePresentation(original));
  for (const [name, data] of Object.entries(parts)) {
    if (!/^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name)) continue;
    parts[name] = strToU8(
      strFromU8(data).replace(
        /<a:bodyPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:bodyPr>)/g,
        name.startsWith('ppt/slides/')
          ? '<a:bodyPr/>'
          : '<a:bodyPr><a:normAutofit fontScale="65000" lnSpcReduction="10000"/></a:bodyPr>',
      ),
    );
  }
  const p = await pptx.loadPresentation(zipSync(parts));
  const shape = pptx
    .getSlideShapes(pptx.getSlides(p)[0])
    .find((s) => pptx.getShapePlaceholderType(s) === 'ctrTitle');
  const model = () => editorModel(p).slides[0].shapes.find((s) => s.id === pptx.getShapeId(shape));
  assert.deepEqual(model().autoFitParams, { fontScale: 0.65, lnSpcReduction: 0.1 });
  pptx.setShapeTextAutoFit(shape, 'normal');
  assert.deepEqual(model().autoFitParams, { fontScale: 1, lnSpcReduction: 0 });
  for (const mode of ['none', 'shape']) {
    pptx.setShapeTextAutoFit(shape, mode);
    assert.equal(model().autoFitParams, null);
  }
});
test('autofit settings persist, replay, undo and detect source changes', async () => {
  for (const autoFit of ['none', 'normal', 'shape']) {
    const p = await fixture();
    const base = await pptx.savePresentation(p);
    const before = editorModel(p);
    const ids = [before.slides[0].shapes[0].id];
    const h = emptyHistory();
    h.entries.push(recordEdit(p, { type: 'update', slide: 0, ids, changes: { autoFit } }));
    h.cursor = 1;
    const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
    assert.equal(editorModel(saved).slides[0].shapes[0].autoFit, autoFit);
    assert.equal(
      pptx.getShapeTextAutoFit(pptx.getSlideShapes(pptx.getSlides(saved)[0])[0]),
      autoFit,
    );
    const replayed = await pptx.loadPresentation(base);
    replayEdits(replayed, h);
    assert.deepEqual(editorModel(replayed), editorModel(saved));
    const conflict = await pptx.loadPresentation(base);
    const shape = pptx.getSlideShapes(pptx.getSlides(conflict)[0])[0];
    pptx.setShapeTextAutoFit(
      shape,
      before.slides[0].shapes[0].autoFit === 'none' ? 'normal' : 'none',
    );
    assert.throws(() => replayEdits(conflict, h), /conflict/i);
    h.cursor = 0;
    const undone = await pptx.loadPresentation(base);
    replayEdits(undone, h);
    assert.deepEqual(editorModel(undone), before);
    assert.throws(() =>
      recordEdit(undone, { type: 'update', slide: 0, ids, changes: { autoFit: 'bad' } }),
    );
  }
});
test('text edits and computed autofit scales save and undo together', async () => {
  const p = await fixture();
  const shape = pptx.getSlideShapes(pptx.getSlides(p)[0])[0];
  pptx.setShapeTextAutoFit(shape, 'normal', { fontScale: 0.8, lnSpcReduction: 0.1 });
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const params = { fontScale: 0.45678, lnSpcReduction: 0.1 };
  const h = emptyHistory();
  h.entries.push(
    recordEdit(p, {
      type: 'update',
      slide: 0,
      ids: [before.slides[0].shapes[0].id],
      changes: { textEdits: [{ start: 8, end: 8, text: ' more text' }], autoFitParams: params },
    }),
  );
  h.cursor = 1;
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.equal(editorModel(saved).slides[0].shapes[0].text, 'Original more text');
  assert.deepEqual(editorModel(saved).slides[0].shapes[0].autoFitParams, params);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, h);
  assert.deepEqual(editorModel(replayed), editorModel(saved));
  h.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, h);
  assert.deepEqual(editorModel(undone), before);
  h.cursor = 1;
  pptx.setShapeTextAutoFit(pptx.getSlideShapes(pptx.getSlides(undone)[0])[0], 'normal', {
    fontScale: 0.7,
    lnSpcReduction: 0.1,
  });
  assert.throws(() => replayEdits(undone, h), /conflict/i);
  for (const bad of [
    { fontScale: 0, lnSpcReduction: 0 },
    { fontScale: 2, lnSpcReduction: 0 },
    { fontScale: 0.5, lnSpcReduction: -1 },
  ])
    assert.throws(() => pptx.setShapeTextAutoFit(shape, 'normal', bad));
});
test('shape autofit resizes about its text anchor and saves with the text edit', async () => {
  const p = await fixture();
  const shape = pptx.getSlideShapes(pptx.getSlides(p)[0])[0];
  pptx.setShapeTextAutoFit(shape, 'shape');
  const base = await pptx.savePresentation(p);
  const model = editorModel(p).slides[0].shapes[0];
  const old = model.bounds;
  const box = resizedTextBounds(model, old.w, old.h * 2);
  assert.deepEqual(box, { ...old, h: old.h * 2 });
  const centered = resizedTextBounds(
    { ...model, textFrame: { ...model.textFrame, anchor: 'center' } },
    old.w,
    old.h * 2,
  );
  assert.equal(centered.y, old.y - old.h / 2);
  const horizontallyCentered = resizedTextBounds(
    { ...model, anchorCenter: true },
    old.w * 2,
    old.h,
  );
  assert.equal(horizontallyCentered.x, old.x - old.w / 2);
  assert.equal(horizontallyCentered.y, old.y);
  const rotated = resizedTextBounds({ ...model, rotation: 90 }, old.w, old.h * 2);
  assert.equal(rotated.x, old.x - old.h / 2);
  assert.equal(rotated.y, old.y - old.h / 2);
  const h = emptyHistory();
  h.entries.push(
    recordEdit(p, {
      type: 'update',
      slide: 0,
      ids: [model.id],
      changes: {
        textEdits: [{ start: 8, end: 8, text: '\nMore text' }],
        bounds: box,
      },
    }),
  );
  h.cursor = 1;
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(saved).slides[0].shapes[0].bounds, box);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, h);
  assert.deepEqual(editorModel(replayed), editorModel(saved));
  h.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, h);
  assert.deepEqual(editorModel(undone).slides[0].shapes[0], model);
});
test('formatting and measured autofit replay and undo as one operation', async () => {
  for (const mode of ['normal', 'shape']) {
    const p = await fixture();
    pptx.setShapeTextAutoFit(pptx.getSlideShapes(pptx.getSlides(p)[0])[0], mode);
    const base = await pptx.savePresentation(p);
    const before = editorModel(p);
    const shape = before.slides[0].shapes[0];
    const fit =
      mode === 'normal'
        ? { autoFitParams: { fontScale: 0.6, lnSpcReduction: 0 } }
        : { bounds: { ...shape.bounds, h: shape.bounds.h * 2 } };
    const command = {
      type: 'update',
      slide: 0,
      ids: [shape.id],
      changes: { format: { size: 72 } },
      textFits: [{ id: shape.id, ...fit }],
    };
    const h = emptyHistory();
    h.entries.push(recordEdit(p, command));
    h.cursor = 1;
    const saved = editorModel(await pptx.loadPresentation(await pptx.savePresentation(p)));
    assert.equal(saved.slides[0].shapes[0].format.size, 72);
    for (const [key, value] of Object.entries(fit))
      assert.deepEqual(saved.slides[0].shapes[0][key], value);
    const replayed = await pptx.loadPresentation(base);
    replayEdits(replayed, h);
    assert.deepEqual(editorModel(replayed), saved);
    h.cursor = 0;
    const undone = await pptx.loadPresentation(base);
    replayEdits(undone, h);
    assert.deepEqual(editorModel(undone), before);
    for (const textFits of [
      [{ id: shape.id + 999, ...fit }],
      [
        { id: shape.id, ...fit },
        { id: shape.id, ...fit },
      ],
      [
        {
          id: shape.id,
          ...(mode === 'normal'
            ? { bounds: shape.bounds }
            : { autoFitParams: { fontScale: 0.6, lnSpcReduction: 0 } }),
        },
      ],
    ]) {
      const invalid = await pptx.loadPresentation(base);
      assert.throws(() => recordEdit(invalid, { ...command, textFits }));
    }
  }
});
test('text directions persist through editing history and exported presentation', async () => {
  for (const direction of ['horz', 'vert', 'vert270', 'wordArtVert']) {
    const p = await fixture();
    const base = await pptx.savePresentation(p);
    const before = editorModel(p);
    const h = emptyHistory();
    h.entries.push(
      recordEdit(p, {
        type: 'update',
        slide: 0,
        ids: [before.slides[0].shapes[0].id],
        changes: { direction },
      }),
    );
    h.cursor = 1;
    const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
    assert.equal(editorModel(saved).slides[0].shapes[0].textDirection, direction);
    const replayed = await pptx.loadPresentation(base);
    replayEdits(replayed, h);
    assert.deepEqual(editorModel(replayed), editorModel(saved));
    h.cursor = 0;
    const undone = await pptx.loadPresentation(base);
    replayEdits(undone, h);
    assert.deepEqual(editorModel(undone), before);
  }
});
test('paragraph formatting follows selection boundaries and a collapsed caret through replay', async () => {
  const p = await fixture();
  const shape = pptx.getSlideShapes(pptx.getSlides(p)[0])[0];
  pptx.setShapeText(shape, 'one\ntwo\nthree');
  const base = await pptx.savePresentation(p);
  const id = pptx.getShapeId(shape);
  const history = emptyHistory();
  for (const changes of [
    {
      range: { start: 4, end: 8 },
      align: 'center',
      bullets: 'number',
      lineSpacing: 1.5,
      indent: 1,
    },
    { range: { start: 8, end: 8 }, align: 'right' },
  ]) {
    history.entries.push(recordEdit(p, { type: 'update', slide: 0, ids: [id], changes }));
    history.cursor++;
  }
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  const result = pptx.getSlideShapes(pptx.getSlides(saved)[0])[0];
  assert.deepEqual(
    [0, 1, 2].map((i) => pptx.getParagraphAlignment(result, i)),
    [null, 'ctr', 'r'],
  );
  assert.deepEqual(
    [0, 1, 2].map((i) => pptx.getParagraphLevel(result, i)),
    [0, 1, 0],
  );
  assert.equal(pptx.getParagraphBullet(result, 1), 'number');
  assert.notEqual(pptx.getParagraphBullet(result, 0), 'number');
  assert.notEqual(pptx.getParagraphBullet(result, 2), 'number');
  assert.deepEqual(pptx.getParagraphLineSpacing(result, 1), { kind: 'pct', value: 1.5 });
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed), editorModel(saved));
  assert.throws(
    () =>
      recordEdit(p, {
        type: 'update',
        slide: 0,
        ids: [id],
        changes: { range: { start: 99, end: 99 }, align: 'left' },
      }),
    /Invalid text range/,
  );
});
test('paragraphs shown while typing match saved splits, joins and empty paragraphs', async () => {
  for (const [start, end, text] of [
    [1, 1, '\n'],
    [3, 4, ''],
    [1, 9, 'X\nY'],
    [0, 0, '\n'],
    [13, 13, '\n'],
  ]) {
    const p = await fixture();
    const shape = pptx.getSlideShapes(pptx.getSlides(p)[0])[0];
    pptx.setShapeText(shape, 'one\ntwo\nthree');
    pptx.setParagraphAlignment(shape, 0, 'left');
    pptx.setParagraphAlignment(shape, 1, 'center');
    pptx.setParagraphAlignment(shape, 2, 'right');
    const before = editorModel(p).slides[0].shapes[0].paragraphs;
    const editing = replaceParagraphRange(before, start, end, text);
    pptx.replaceShapeTextRange(shape, start, end, text);
    assert.deepEqual(
      editing,
      editorModel(p).slides[0].shapes[0].paragraphs,
      JSON.stringify([start, end, text]),
    );
  }
});
test('case changes preserve each selected text and its run formatting', async () => {
  const p = await fixture();
  const slide = pptx.getSlides(p)[0];
  const first = pptx.getSlideShapes(slide)[0];
  pptx.setShapeTextFormat(first, { bold: true });
  pptx.addSlideTextBox(slide, { x: 0, y: 0, w: 1000000, h: 1000000, text: 'Other 日本語' });
  const ids = editorModel(p).slides[0].shapes.map((s) => s.id);
  recordEdit(p, { type: 'update', slide: 0, ids, changes: { changeCase: 'upper' } });
  const shapes = editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))).slides[0]
    .shapes;
  assert.deepEqual(
    shapes.map((s) => s.text),
    ['ORIGINAL', 'OTHER 日本語'],
  );
  assert.equal(shapes[0].format.bold, true);
});
test('native aspect locks toggle without geometry changes and survive export, replay and undo', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const before = editorModel(p).slides[0].shapes[0];
  const h = emptyHistory();
  h.entries.push(
    recordEdit(p, {
      type: 'update',
      slide: 0,
      ids: [before.id],
      changes: { aspectRatioLocked: true },
    }),
  );
  h.cursor++;
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.equal(editorModel(saved).slides[0].shapes[0].aspectRatioLocked, true);
  assert.deepEqual(editorModel(saved).slides[0].shapes[0].bounds, before.bounds);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, h);
  assert.deepEqual(editorModel(replayed), editorModel(saved));
  const conflict = await pptx.loadPresentation(base);
  pptx.setShapeAspectRatioLocked(pptx.getSlideShapes(pptx.getSlides(conflict)[0])[0], true);
  assert.throws(() => replayEdits(conflict, h), /conflict/i);
  h.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, h);
  assert.equal(editorModel(undone).slides[0].shapes[0].aspectRatioLocked, false);
  recordEdit(saved, {
    type: 'update',
    slide: 0,
    ids: [before.id],
    changes: { aspectRatioLocked: false },
  });
  const unlocked = await pptx.loadPresentation(await pptx.savePresentation(saved));
  assert.equal(editorModel(unlocked).slides[0].shapes[0].aspectRatioLocked, false);
  assert.throws(
    () =>
      recordEdit(unlocked, {
        type: 'update',
        slide: 0,
        ids: [before.id],
        changes: { aspectRatioLocked: 'yes' },
      }),
    /Invalid aspect ratio lock/,
  );
});
test('picture bytes, name and aspect ratio survive export and replay', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const h = emptyHistory();
  h.entries.push(
    recordEdit(p, {
      type: 'insert',
      slide: 0,
      preset: 'picture',
      image: {
        name: 'sample.png',
        base64:
          'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jp1sAAAAASUVORK5CYII=',
      },
      changes: { bounds: { x: 0, y: 0, w: 2000000, h: 1000000 } },
    }),
  );
  h.cursor++;
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  const shape = editorModel(saved).slides[0].shapes.at(-1);
  assert.equal(shape.kind, 'picture');
  assert.equal(shape.aspectRatioLocked, true);
  assert.equal(editorModel(saved).slides[0].shapes[0].aspectRatioLocked, false);
  assert.equal(shape.name, 'sample.png');
  assert.deepEqual(shape.bounds, { x: 500000, y: 0, w: 1000000, h: 1000000 });
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, h);
  assert.deepEqual(editorModel(replayed), editorModel(saved));
  assert.throws(
    () =>
      recordEdit(replayed, {
        type: 'insert',
        slide: 0,
        preset: 'picture',
        image: { name: 'bad.png', base64: 'YWJj' },
      }),
    /could not detect image format/,
  );
});
test('visual edits replay across serialization, preserve source, and support undo/redo', async () => {
  const base = await fixture(),
    bytes = await pptx.savePresentation(base);
  let p = await pptx.loadPresentation(bytes);
  const h = emptyHistory(),
    id = editorModel(p).slides[0].shapes[0].id;
  for (const command of [
    {
      type: 'update',
      slide: 0,
      ids: [id],
      changes: { text: 'Edited 日本語', format: { size: 28, bold: true } },
    },
    {
      type: 'update',
      slide: 0,
      ids: [id],
      changes: { bounds: { x: 1200000, y: 1000000, w: 3000000, h: 700000 }, rotation: 15 },
    },
    { type: 'duplicate', slide: 0, ids: [id] },
    { type: 'notes', slide: 0, text: 'Speaker notes' },
    { type: 'slide-duplicate', slide: 0 },
    { type: 'slide-move', slide: 1, to: 0 },
    { type: 'slide-hidden', slide: 1, hidden: true },
  ]) {
    h.entries.push(recordEdit(p, command));
    h.cursor++;
    p = await pptx.loadPresentation(await pptx.savePresentation(p));
    const rebuilt = await pptx.loadPresentation(bytes);
    replayEdits(rebuilt, h);
    assert.deepEqual(editorModel(rebuilt), editorModel(p));
  }
  const undone = await pptx.loadPresentation(bytes);
  replayEdits(undone, { ...h, cursor: 0 });
  assert.equal(pptx.getSlideText(pptx.getSlides(undone)[0]), 'Original');
  const redone = await pptx.loadPresentation(bytes);
  replayEdits(redone, h);
  assert.equal(editorModel(redone).slides[1].hidden, true);
  assert.equal(editorModel(base).slides[0].shapes[0].text, 'Original');
});
test('replay refuses conflicting source changes; malformed operations do not reach persistence', async () => {
  const p = await fixture(),
    h = emptyHistory();
  const id = editorModel(p).slides[0].shapes[0].id;
  h.entries.push(recordEdit(p, { type: 'delete', slide: 0, ids: [id] }));
  h.cursor++;
  const changed = await fixture();
  pptx.setShapeText(pptx.getSlideShapes(pptx.getSlides(changed)[0])[0], 'Source changed');
  assert.throws(() => replayEdits(changed, h), /conflicts/);
  assert.throws(
    () =>
      recordEdit(changed, {
        type: 'update',
        slide: 0,
        ids: [id],
        changes: { bounds: { x: 0, y: 0, w: -1, h: 50 } },
      }),
    /Invalid/,
  );
  assert.throws(() => recordEdit(changed, { type: 'slide-move', slide: 0, to: 0.5 }), /Invalid/);
});

test('new slides use the selected deck layout and preserve placeholders through replay', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const layout = editorModel(p).layouts.find((item) => item.type === 'obj');
  assert.ok(layout);
  const history = emptyHistory();
  history.entries.push(recordEdit(p, { type: 'slide-add', slide: 0, layout: layout.key }));
  history.cursor++;
  const inserted = editorModel(p).slides[1];
  assert.equal(inserted.layout, layout.key);
  assert.ok(inserted.shapes.length >= 2);
  assert.ok(inserted.shapes.every((s) => s.text === ''));
  assert.ok(inserted.shapes.some((s) => s.placeholder === 'title'));
  assert.ok(inserted.shapes.some((s) => s.bounds?.w > 0));
  const roundtrip = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(roundtrip), editorModel(p));
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed), editorModel(p));
  assert.throws(
    () => recordEdit(p, { type: 'slide-add', slide: 0, layout: '/missing.xml' }),
    /layout/,
  );
  assert.equal(pptx.getSlides(p).length, 2);
});

test('layout application and reset retain text and ordinary objects through replay', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const history = emptyHistory();
  const apply = (command) => {
    history.entries.push(recordEdit(p, command));
    history.cursor++;
  };
  const layout = editorModel(p).layouts.find((item) => item.type === 'obj');
  const original = editorModel(p).slides[0].shapes[0];
  apply({ type: 'slide-layout', slide: 0, layout: layout.key });
  let slide = editorModel(p).slides[0];
  assert.deepEqual(
    slide.shapes.find((s) => s.id === original.id),
    original,
  );
  const title = slide.shapes.find((s) => s.placeholder === 'title');
  assert.ok(title?.bounds);
  apply({
    type: 'update',
    slide: 0,
    ids: [title.id],
    changes: {
      text: 'Preserved title',
      bounds: { ...title.bounds, x: title.bounds.x + 100000 },
      format: { size: 66, bold: true },
    },
  });
  const body = slide.shapes.find((s) => s.placeholder === 'obj');
  apply({ type: 'delete', slide: 0, ids: [body.id] });
  apply({ type: 'slide-reset', slide: 0 });
  slide = editorModel(p).slides[0];
  const reset = slide.shapes.find((s) => s.id === title.id);
  assert.equal(reset.text, 'Preserved title');
  assert.deepEqual(reset.bounds, title.bounds);
  assert.notEqual(reset.format.size, 66);
  assert.ok(slide.shapes.some((s) => s.placeholder === 'obj'));
  assert.deepEqual(
    slide.shapes.find((s) => s.id === original.id),
    original,
  );
  const titleLayout = editorModel(p).layouts.find((item) => item.type === 'title');
  apply({ type: 'slide-layout', slide: 0, layout: titleLayout.key });
  assert.equal(
    editorModel(p).slides[0].shapes.find((s) => s.id === title.id).text,
    'Preserved title',
  );
  assert.equal(
    editorModel(p).slides[0].shapes.find((s) => s.id === title.id).placeholder,
    'ctrTitle',
  );
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed), editorModel(p));
  assert.deepEqual(
    editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))),
    editorModel(p),
  );
});

test('group selection exposes only the group and ungroup restores the children', async () => {
  const p = await fixture();
  const id = editorModel(p).slides[0].shapes[0].id;
  recordEdit(p, { type: 'duplicate', slide: 0, ids: [id] });
  const ids = editorModel(p).slides[0].shapes.map((shape) => shape.id);
  recordEdit(p, { type: 'group', slide: 0, ids });
  const grouped = editorModel(p).slides[0].shapes;
  assert.equal(grouped.length, 1);
  assert.equal(grouped[0].kind, 'group');
  recordEdit(p, { type: 'ungroup', slide: 0, ids: [grouped[0].id] });
  assert.equal(editorModel(p).slides[0].shapes.length, 2);
});

test('group descendants retain hierarchy, local frames and slide transforms through edits', async () => {
  const p = await fixture();
  const slide = pptx.getSlides(p)[0];
  const first = pptx.getSlideShapes(slide)[0];
  const firstId = pptx.getShapeId(first);
  recordEdit(p, { type: 'duplicate', slide: 0, ids: [firstId] });
  recordEdit(p, { type: 'group', slide: 0, ids: editorModel(p).slides[0].shapes.map((s) => s.id) });
  let root = editorModel(p).slides[0].shapes[0];
  const group = pptx.getSlideShapes(slide).find((s) => pptx.getShapeId(s) === root.id);
  pptx.setShapeRotation(group, 90);
  root = editorModel(p).slides[0].shapes[0];
  const svg = renderSlideToSvg(p, slide);
  for (const child of root.children) assert.ok(svg.includes(`data-pptx-shape-id="${child.id}"`));
  assert.equal(root.children.length, 2);
  assert.equal(root.parentTransform, undefined);
  const child = root.children.find((s) => s.id === firstId);
  assert.equal(child.text, 'Original');
  assert.equal(child.textable, true);
  const [a, b, c, d, e, f] = child.parentTransform;
  assert.ok(Math.abs(a) < 1e-9);
  assert.ok(Math.abs(b - 1) < 1e-9);
  assert.ok(Math.abs(c + 1) < 1e-9);
  assert.ok(Math.abs(d) < 1e-9);
  const center = { x: root.bounds.x + root.bounds.w / 2, y: root.bounds.y + root.bounds.h / 2 };
  assert.ok(Math.abs(a * center.x + c * center.y + e - center.x) < 0.01);
  assert.ok(Math.abs(b * center.x + d * center.y + f - center.y) < 0.01);
  // A second group level must not flatten the first group's children.
  pptx.addSlideTextBox(slide, { x: 0, y: 0, w: 914400, h: 914400, text: 'Outside' });
  recordEdit(p, { type: 'group', slide: 0, ids: editorModel(p).slides[0].shapes.map((s) => s.id) });
  const before = editorModel(p);
  const nested = before.slides[0].shapes[0].children.find((s) => s.id === root.id);
  assert.equal(nested.children.length, 2);
  assert.deepEqual(nested.children[0].parentTransform, root.children[0].parentTransform);
  const bytes = await pptx.savePresentation(p);
  const history = emptyHistory();
  history.entries.push(
    recordEdit(p, {
      type: 'update',
      slide: 0,
      ids: [firstId],
      changes: { text: 'Child edited', format: { bold: true } },
    }),
  );
  history.cursor = 1;
  const after = editorModel(p);
  const reload = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(reload), after);
  const replay = await pptx.loadPresentation(bytes);
  replayEdits(replay, history);
  assert.deepEqual(editorModel(replay), after);
  history.cursor = 0;
  const undo = await pptx.loadPresentation(bytes);
  replayEdits(undo, history);
  assert.deepEqual(editorModel(undo), before);
});

test('font and paragraph commands survive PPTX serialization and journal replay', async () => {
  const p = await fixture(),
    bytes = await pptx.savePresentation(p),
    history = emptyHistory();
  const id = editorModel(p).slides[0].shapes[0].id;
  for (const changes of [
    { format: { strike: true, baseline: 0.3, spc: 300, highlight: 'FFFF00' } },
    {
      bullets: 'number',
      anchor: 'bottom',
      direction: 'vert270',
      columns: 2,
      lineSpacing: 1.5,
      indent: 1,
      align: 'distribute',
    },
  ]) {
    history.entries.push(recordEdit(p, { type: 'update', slide: 0, ids: [id], changes }));
    history.cursor++;
  }
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  const shape = pptx.getSlideShapes(pptx.getSlides(saved)[0])[0];
  const format = pptx.getShapeRunFormat(shape, 0, 0);
  assert.equal(format.strike, true);
  assert.equal(format.baseline, 0.3);
  assert.equal(format.spc, 300);
  assert.equal(format.highlight, '#FFFF00');
  assert.equal(pptx.getParagraphBullet(shape, 0), 'number');
  assert.equal(pptx.getParagraphLevel(shape, 0), 1);
  assert.deepEqual(pptx.getParagraphLineSpacing(shape, 0), { kind: 'pct', value: 1.5 });
  assert.equal(pptx.getShapeTextAnchor(shape), 'bottom');
  assert.equal(pptx.getShapeTextDirection(shape), 'vert270');
  assert.equal(pptx.getShapeTextColumns(shape).count, 2);
  const replayed = await pptx.loadPresentation(bytes);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed), editorModel(saved));
});

test('insertion formatting affects new text only and replays from the journal', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const id = editorModel(p).slides[0].shapes[0].id;
  const history = emptyHistory();
  history.entries.push(
    recordEdit(p, {
      type: 'update',
      slide: 0,
      ids: [id],
      changes: {
        textEdits: [
          { start: 3, end: 3, text: 'New', format: { italic: true, color: '#FF0000' } },
          { start: 6, end: 6, text: '!', format: { italic: false, color: '#0000FF' } },
        ],
      },
    }),
  );
  history.cursor++;
  const loaded = await pptx.loadPresentation(base);
  replayEdits(loaded, history);
  const result = editorModel(loaded).slides[0].shapes[0];
  assert.equal(result.text, 'OriNew!ginal');
  assert.deepEqual(
    result.runs.map((r) => [r.start, r.end]),
    [
      [0, 3],
      [3, 6],
      [6, 7],
      [7, 12],
    ],
  );
  assert.equal(result.runs[1].format.italic, true);
  assert.equal(result.runs[1].format.color, '#FF0000');
  assert.equal(result.runs[2].format.italic, false);
  assert.equal(result.runs[2].format.color, '#0000FF');
  assert.notEqual(result.runs[3].format.italic, true);
  assert.notEqual(result.runs[3].format.color, '#0000FF');
});

test('categorized shape presets render and survive export and journal replay', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const history = emptyHistory();
  const presets = shapeCategories.flatMap((category) =>
    category.shapes.map((shape) => shape.preset),
  );
  assert.equal(new Set(presets).size, presets.length);
  for (const preset of presets) {
    const path = getPresetShapePath(preset);
    assert.ok(path && !/NaN|undefined|Infinity/.test(path), preset);
    history.entries.push(
      recordEdit(p, {
        type: 'insert',
        slide: 0,
        preset,
        changes: { bounds: { x: 0, y: 0, w: 2000000, h: 1000000 } },
      }),
    );
    history.cursor++;
  }
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(
    pptx.getSlideShapes(pptx.getSlides(saved)[0]).slice(1).map(pptx.getShapePreset),
    presets,
  );
  const buttons = editorModel(saved).slides[0].shapes.filter((shape) =>
    shape.preset?.startsWith('actionButton'),
  );
  assert.equal(buttons.length, 12);
  const destinations = {
    actionButtonBackPrevious: 'prevSlide',
    actionButtonForwardNext: 'nextSlide',
    actionButtonBeginning: 'firstSlide',
    actionButtonEnd: 'lastSlide',
    actionButtonHome: 'firstSlide',
    actionButtonReturn: 'lastSlideViewed',
  };
  for (const button of buttons) {
    assert.deepEqual(
      button.link?.action ?? null,
      destinations[button.preset] ? { kind: destinations[button.preset] } : null,
    );
  }
  const svg = renderSlideToSvg(saved, pptx.getSlides(saved)[0]);
  assert.ok(!svg.includes('data-pptx-preset='), 'Every gallery shape has a preview geometry');
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed), editorModel(saved));
  for (const preset of ['__proto__', 'constructor', 'unrecognized']) {
    assert.equal(getPresetShapePath(preset), null);
    assert.throws(
      () => recordEdit(p, { type: 'insert', slide: 0, preset }),
      /Invalid shape preset/,
    );
  }
});

test('line tools preserve routing, reversed endpoints and arrows through export, movement and replay', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const history = emptyHistory();
  for (const [index, line] of lineTools.entries()) {
    const from = { x: 2000000, y: 1000000 },
      toPoint = { x: 0, y: index % 2 === 0 ? 1000000 : 0 };
    history.entries.push(
      recordEdit(p, { type: 'insert', slide: 0, preset: line.preset, from, toPoint }),
    );
    history.cursor++;
  }
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  const shapes = pptx.getSlideShapes(pptx.getSlides(saved)[0]).slice(1);
  for (const [index, line] of lineTools.entries()) {
    const shape = shapes[index];
    assert.equal(pptx.getShapeKind(shape), 'connector');
    assert.equal(pptx.getShapePreset(shape), line.geometry);
    assert.deepEqual(pptx.getShapeFlip(shape), { horizontal: true, vertical: index % 2 !== 0 });
    assert.equal(
      pptx.getShapeStrokeArrow(shape, 'tail')?.type ?? 'none',
      line.arrows > 0 ? 'triangle' : 'none',
    );
    assert.equal(
      pptx.getShapeStrokeArrow(shape, 'head')?.type ?? 'none',
      line.arrows > 1 ? 'triangle' : 'none',
    );
  }
  const line = editorModel(p).slides[0].shapes[1];
  history.entries.push(
    recordEdit(p, {
      type: 'update',
      slide: 0,
      ids: [line.id],
      positions: [{ id: line.id, bounds: { ...line.bounds, x: 500000 } }],
    }),
  );
  history.cursor++;
  history.entries.push(
    recordEdit(p, {
      type: 'update',
      slide: 0,
      ids: [line.id],
      changes: {
        bounds: { x: 2500000, y: 1000000, w: 900000, h: 600000 },
        flip: { horizontal: false, vertical: false },
      },
    }),
  );
  history.cursor++;
  const exported = await pptx.loadPresentation(await pptx.savePresentation(p));
  const edited = editorModel(exported).slides[0].shapes[1];
  assert.deepEqual(edited.flip, { horizontal: false, vertical: false });
  assert.deepEqual(edited.bounds, { x: 2500000, y: 1000000, w: 900000, h: 600000 });
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed), editorModel(p));
  assert.throws(() => recordEdit(p, { type: 'insert', slide: 0, preset: 'line' }), /endpoints/);
});

test('connectors retain attachment across transformed group coordinate systems', async () => {
  const p = await fixture();
  const slide = pptx.getSlides(p)[0];
  const target = pptx.getSlideShapes(slide)[0];
  const targetId = pptx.getShapeId(target);
  const edit = (command) => recordEdit(p, { slide: 0, ...command });
  edit({ type: 'duplicate', ids: [targetId] });
  const otherId = editorModel(p).slides[0].shapes[1].id;
  edit({
    type: 'insert',
    preset: 'lineArrow',
    from: { x: 0, y: 0 },
    toPoint: { x: 6000000, y: 4000000 },
    connections: { start: { shapeId: targetId, siteIndex: 3 } },
  });
  const lineId = editorModel(p).slides[0].shapes[2].id;
  edit({ type: 'group', ids: [targetId, otherId] });
  const group = editorModel(p).slides[0].shapes.find((s) => s.kind === 'group');
  edit({
    type: 'update',
    ids: [group.id],
    changes: {
      bounds: { x: 2000000, y: 1000000, w: 5000000, h: 3000000 },
      rotation: 37,
      flip: { horizontal: true, vertical: false },
    },
  });
  const { shapePoint, connectorEndpoints } = await import('../src/connector-geometry.ts');
  const find = (id) => pptx.getSlideShapes(slide).find((s) => pptx.getShapeId(s) === id);
  const frame = (s) => ({
    bounds: pptx.getShapeBoundsResolved(p, s),
    rotation: pptx.getShapeRotation(s),
    flip: pptx.getShapeFlip(s),
  });
  const targetPoint = shapePoint(
    frame(find(targetId)),
    pptx.getShapeConnectionSites(find(targetId))[3],
  );
  const transform = pptx.getGroupTransform(find(group.id));
  const expected = shapePoint(frame(find(group.id)), {
    x: ((targetPoint.x - transform.inner.x) * transform.outer.w) / transform.inner.w,
    y: ((targetPoint.y - transform.inner.y) * transform.outer.h) / transform.inner.h,
  });
  const actual = connectorEndpoints(frame(find(lineId))).start;
  assert.ok(Math.hypot(actual.x - expected.x, actual.y - expected.y) < 2);
  // Moving a parent containing both the target group and connector must not
  // reinterpret either child's coordinates as slide coordinates.
  edit({ type: 'group', ids: [group.id, lineId] });
  const outer = editorModel(p).slides[0].shapes[0];
  const before = connectorEndpoints(frame(find(lineId)));
  edit({
    type: 'update',
    ids: [outer.id],
    changes: {
      bounds: { ...outer.bounds, x: 3000000, y: 2000000, w: outer.bounds.w * 1.3 },
      rotation: 23,
    },
  });
  const after = connectorEndpoints(frame(find(lineId)));
  assert.ok(Math.hypot(after.start.x - before.start.x, after.start.y - before.start.y) < 2);
  assert.deepEqual(after.end, before.end);
  assert.deepEqual(pptx.getShapeConnection(find(lineId), 'start'), {
    shapeId: targetId,
    siteIndex: 3,
  });
  edit({ type: 'duplicate', ids: [outer.id] });
  const copiedOuter = editorModel(p).slides[0].shapes[1];
  const copiedChildren = pptx.getGroupChildren(find(copiedOuter.id));
  const copiedLine = copiedChildren.find((s) => pptx.getShapeKind(s) === 'connector');
  const copiedGroup = copiedChildren.find((s) => pptx.getShapeKind(s) === 'group');
  const copiedTarget = pptx.getGroupChildren(copiedGroup)[0];
  assert.deepEqual(pptx.getShapeConnection(copiedLine, 'start'), {
    shapeId: pptx.getShapeId(copiedTarget),
    siteIndex: 3,
  });
  assert.notEqual(pptx.getShapeId(copiedTarget), targetId);
  const copiedPoint = connectorEndpoints(frame(copiedLine)).start;
  assert.ok(Math.hypot(copiedPoint.x - after.start.x, copiedPoint.y - after.start.y) < 2);
});

test('attached connectors follow move/resize/rotation, detach and remap copied selections', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const history = emptyHistory();
  const apply = (command) => {
    history.entries.push(recordEdit(p, command));
    history.cursor++;
  };
  const model = () => editorModel(p).slides[0].shapes;
  const target = model()[0];
  const attachment = { shapeId: target.id, siteIndex: 3 };
  apply({
    type: 'insert',
    slide: 0,
    preset: 'lineArrow',
    from: { x: 0, y: 0 },
    toPoint: { x: 3000000, y: 3000000 },
    connections: { start: attachment },
  });
  const lineId = model()[1].id;
  const { connectorEndpoints } = await import('../src/connector-geometry.ts');
  const check = () => {
    const [target, line] = model();
    const point = connectorEndpoints(line).start;
    const site = target.connectionSites[3];
    assert.ok(Math.abs(point.x - site.x) < 2 && Math.abs(point.y - site.y) < 2);
    assert.deepEqual(line.connections.start, attachment);
  };
  check();
  apply({
    type: 'update',
    slide: 0,
    ids: [target.id],
    changes: {
      bounds: { x: 1000000, y: 1200000, w: 2000000, h: 1500000 },
      rotation: 37,
      flip: { horizontal: true, vertical: false },
    },
  });
  check();
  const lineBefore = model()[1];
  apply({ type: 'duplicate', slide: 0, ids: [lineId, target.id] });
  const copies = model().slice(2);
  assert.equal(copies[0].connections.start.shapeId, copies[1].id);
  assert.notEqual(copies[0].connections.start.shapeId, target.id);
  assert.ok(Math.abs(connectorEndpoints(copies[0]).start.x - copies[1].connectionSites[3].x) < 2);
  apply({ type: 'delete', slide: 0, ids: copies.map((s) => s.id) });
  apply({ type: 'duplicate', slide: 0, ids: [lineId] });
  assert.equal(model()[2].connections.start, null);
  apply({ type: 'delete', slide: 0, ids: [model()[2].id] });
  const exported = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(exported), editorModel(p));
  apply({
    type: 'update',
    slide: 0,
    ids: [lineId],
    changes: { connections: { start: null }, bounds: { ...lineBefore.bounds, x: 0 } },
  });
  assert.equal(model()[1].connections.start, null);
  apply({
    type: 'update',
    slide: 0,
    ids: [lineId],
    changes: { connections: { start: attachment } },
  });
  check();
  apply({ type: 'delete', slide: 0, ids: [target.id] });
  assert.equal(model()[0].connections.start, null);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed), editorModel(p));
  history.cursor--;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, history);
  assert.deepEqual(editorModel(undone).slides[0].shapes[1].connections.start, attachment);
});

test('sections persist through slide edits, replay, undo and PPTX serialization', async () => {
  const p = await fixture();
  for (let i = 0; i < 3; i++) pptx.addBlankSlide(p);
  const base = await pptx.savePresentation(p);
  const history = emptyHistory();
  const edit = (command) => {
    history.entries.push(recordEdit(p, command));
    history.cursor++;
  };
  edit({ type: 'section-add', slide: 2, text: 'Second part' });
  assert.deepEqual(
    editorModel(p).sections.map((s) => [s.name, s.slides.length]),
    [
      ['Default Section', 2],
      ['Second part', 2],
    ],
  );
  edit({ type: 'section-rename', slide: 2, section: 1, text: '説明 <script>' });
  edit({ type: 'slide-duplicate', slide: 2 });
  assert.equal(editorModel(p).sections[1].slides.length, 3);
  edit({ type: 'slide-add', slide: 0 });
  assert.equal(editorModel(p).sections[0].slides.length, 3);
  edit({ type: 'slide-move', slide: 0, to: 5 });
  assert.deepEqual(
    editorModel(p).sections.map((s) => s.slides.length),
    [2, 4],
  );
  edit({ type: 'slide-delete', slide: 5 });
  assert.deepEqual(
    editorModel(p).sections.map((s) => s.slides.length),
    [2, 3],
  );
  const expected = editorModel(p);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed), expected);
  assert.deepEqual(
    editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))),
    expected,
  );
  edit({ type: 'section-remove', slide: 2, section: 1 });
  assert.equal(editorModel(p).sections.length, 1);
  assert.equal(editorModel(p).sections[0].slides.length, 5);
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...history, cursor: history.cursor - 1 });
  assert.deepEqual(editorModel(undone), expected);
  edit({ type: 'section-remove-all', slide: 0 });
  assert.equal(editorModel(p).sections.length, 0);
  assert.equal(editorModel(p).slides.length, 5);
  const conflict = await pptx.loadPresentation(base);
  pptx.setSlideSections(conflict, [{ name: 'External section', slides: pptx.getSlides(conflict) }]);
  assert.throws(() => replayEdits(conflict, history), /conflict/i);
});

test('section moves and section-and-slide deletion preserve slide identity and support undo', async () => {
  const p = await fixture();
  for (let i = 0; i < 5; i++) {
    const slide = pptx.addBlankSlide(p);
    pptx.addSlideTextBox(slide, {
      x: 0,
      y: 0,
      w: pptx.inches(2),
      h: pptx.inches(1),
      text: 'Slide ' + (i + 2),
    });
  }
  const slides = pptx.getSlides(p);
  pptx.setSlideSections(p, [
    { name: 'First', slides: slides.slice(0, 2) },
    { name: 'Empty', slides: [] },
    { name: 'Middle', slides: slides.slice(2, 4) },
    { name: 'Last', slides: slides.slice(4) },
  ]);
  const base = await pptx.savePresentation(p);
  const originalKeys = editorModel(p).slides.map((s) => s.key);
  const history = emptyHistory();
  const edit = (command) => {
    history.entries.push(recordEdit(p, command));
    history.cursor++;
  };
  edit({ type: 'section-move', slide: 0, section: 3, to: 0 });
  assert.deepEqual(
    editorModel(p).sections.map((s) => s.name),
    ['Last', 'First', 'Empty', 'Middle'],
  );
  assert.deepEqual(
    editorModel(p).slides.map((s) => s.key),
    [...originalKeys.slice(4), ...originalKeys.slice(0, 4)],
  );
  edit({ type: 'section-move', slide: 0, section: 2, to: 3 });
  assert.deepEqual(
    editorModel(p).sections.map((s) => s.name),
    ['Last', 'First', 'Middle', 'Empty'],
  );
  const beforeDelete = editorModel(p);
  edit({ type: 'section-delete', slide: 0, section: 1 });
  assert.deepEqual(
    editorModel(p).slides.map((s) => s.key),
    [...originalKeys.slice(4), ...originalKeys.slice(2, 4)],
  );
  assert.deepEqual(
    editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))),
    editorModel(p),
  );
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...history, cursor: history.cursor - 1 });
  assert.deepEqual(editorModel(undone), beforeDelete);
  const redone = await pptx.loadPresentation(base);
  replayEdits(redone, history);
  assert.deepEqual(editorModel(redone), editorModel(p));
  const unchanged = editorModel(p);
  assert.throws(
    () => recordEdit(p, { type: 'section-move', slide: 0, section: 0, to: 9 }),
    /Invalid section position/,
  );
  assert.deepEqual(editorModel(p), unchanged);
  edit({ type: 'section-delete', slide: 0, section: 0 });
  edit({ type: 'section-delete', slide: 0, section: 0 });
  assert.equal(editorModel(p).slides.length, 0);
  assert.deepEqual(editorModel(p).sections, [{ name: 'Empty', slides: [] }]);
  edit({ type: 'section-rename', slide: 0, section: 0, text: 'Still editable' });
  edit({ type: 'section-delete', slide: 0, section: 0 });
  assert.deepEqual(editorModel(p).sections, []);
});

test('relative rotation preserves individual angles, geometry and replay', async () => {
  const p = await fixture();
  const slide = pptx.getSlides(p)[0];
  const first = pptx.getSlideShapes(slide)[0];
  pptx.setShapeRotation(first, 350);
  recordEdit(p, { type: 'duplicate', slide: 0, ids: [editorModel(p).slides[0].shapes[0].id] });
  pptx.setShapeRotation(pptx.getSlideShapes(slide)[1], 25);
  const base = await pptx.savePresentation(p);
  const original = editorModel(p).slides[0].shapes;
  const ids = original.map((shape) => shape.id);
  const h = emptyHistory();
  h.entries.push(recordEdit(p, { type: 'update', slide: 0, ids, changes: { rotationDelta: 15 } }));
  h.cursor++;
  assert.deepEqual(
    editorModel(p).slides[0].shapes.map((shape) => shape.rotation),
    [5, 40],
  );
  assert.deepEqual(
    editorModel(p).slides[0].shapes.map((shape) => shape.bounds),
    original.map((shape) => shape.bounds),
  );
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, h);
  assert.deepEqual(editorModel(replayed), editorModel(saved));
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...h, cursor: 0 });
  assert.deepEqual(
    editorModel(undone).slides[0].shapes.map((shape) => shape.rotation),
    [350, 25],
  );
  recordEdit(p, { type: 'update', slide: 0, ids, changes: { rotationDelta: -15 } });
  assert.deepEqual(
    editorModel(p).slides[0].shapes.map((shape) => shape.rotation),
    [350, 25],
  );
  assert.throws(() =>
    recordEdit(p, { type: 'update', slide: 0, ids, changes: { rotationDelta: NaN } }),
  );
  assert.throws(
    () =>
      recordEdit(p, {
        type: 'update',
        slide: 0,
        ids,
        changes: { rotationDelta: 15, rotation: 30 },
      }),
    /Cannot combine/,
  );
});

test('centered text anchors save, replay and undo independently of paragraph alignment', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const id = editorModel(p).slides[0].shapes[0].id;
  const h = emptyHistory();
  h.entries.push(
    recordEdit(p, {
      type: 'update',
      slide: 0,
      ids: [id],
      changes: { anchor: 'bottom', anchorCenter: true },
    }),
  );
  h.cursor = 1;
  const current = editorModel(p).slides[0].shapes[0];
  assert.equal(current.anchorCenter, true);
  assert.equal(current.textFrame.anchor, 'bottom');
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.equal(editorModel(saved).slides[0].shapes[0].anchorCenter, true);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, h);
  assert.equal(editorModel(replayed).slides[0].shapes[0].anchorCenter, true);
  const conflict = await pptx.loadPresentation(base);
  pptx.setShapeTextAnchorCenter(pptx.getSlideShapes(pptx.getSlides(conflict)[0])[0], true);
  assert.throws(() => replayEdits(conflict, h), /conflict/i);
  h.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, h);
  assert.equal(editorModel(undone).slides[0].shapes[0].anchorCenter, false);
});

test('drawing guides persist, replay, undo, and detect source conflicts', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const h = emptyHistory();
  h.entries.push(
    recordEdit(p, {
      type: 'guides',
      slide: 0,
      guides: [{ id: 9, axis: 'y', offset: 914400, color: '#2873c4' }],
    }),
  );
  h.cursor = 1;
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(saved).guides, [
    { id: 9, axis: 'y', offset: 914400, color: '#2873c4' },
  ]);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, h);
  assert.deepEqual(editorModel(replayed), editorModel(saved));
  const conflict = await pptx.loadPresentation(base);
  pptx.setDrawingGuides(conflict, []);
  assert.throws(() => replayEdits(conflict, h), /conflict/i);
  h.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, h);
  assert.deepEqual(editorModel(undone), before);
});

test('grid snapping saves per document, replays, undoes and detects conflicts', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const h = emptyHistory();
  h.entries.push(recordEdit(p, { type: 'grid-snap', slide: 0, snapToGrid: true }));
  h.cursor = 1;
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.equal(editorModel(saved).snapToGrid, true);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, h);
  assert.deepEqual(editorModel(replayed), editorModel(saved));
  const conflict = await pptx.loadPresentation(base);
  pptx.setSnapToGrid(conflict, false);
  assert.throws(() => replayEdits(conflict, h), /conflict/i);
  h.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, h);
  assert.deepEqual(editorModel(undone), before);
});

test('guide visibility saves, replays and restores through history', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const h = emptyHistory();
  assert.equal(editorModel(p).guidesVisible, false);
  h.entries.push(recordEdit(p, { type: 'guides-visible', slide: 0, guidesVisible: true }));
  h.cursor = 1;
  assert.equal(
    editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))).guidesVisible,
    true,
  );
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, h);
  assert.equal(editorModel(replayed).guidesVisible, true);
  assert.throws(() => replayEdits(replayed, h), /conflict/i);
  h.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, h);
  assert.equal(editorModel(undone).guidesVisible, false);
});

test('canvas size persists, replays, undoes and detects changed source dimensions', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const h = emptyHistory();
  const size = { width: 6858000, height: 9144000, type: 'custom' };
  h.entries.push(recordEdit(p, { type: 'slide-size', slide: 0, size }));
  h.cursor = 1;
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(pptx.getSlideSize(saved), size);
  assert.deepEqual(editorModel(saved).slides, before.slides);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, h);
  assert.deepEqual(editorModel(replayed), editorModel(saved));
  const conflict = await pptx.loadPresentation(base);
  pptx.setSlideSize(conflict, pptx.SLIDE_SIZE_4_3);
  assert.throws(() => replayEdits(conflict, h), /conflict/i);
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...h, cursor: 0 });
  assert.deepEqual(editorModel(undone), before);
  replayEdits(undone, h);
  assert.deepEqual(editorModel(undone), editorModel(saved));
  for (const invalid of [
    undefined,
    {},
    { ...size, width: '6858000' },
    { ...size, height: NaN },
    { ...size, height: 0 },
    { ...size, width: Infinity },
    { ...size, type: 'invalid' },
  ]) {
    assert.throws(() => recordEdit(p, { type: 'slide-size', slide: 0, size: invalid }));
    assert.deepEqual(pptx.getSlideSize(p), size);
  }
  const empty = pptx.createPresentation();
  recordEdit(empty, { type: 'slide-size', slide: 0, size });
  assert.deepEqual(pptx.getSlideSize(empty), size);
});

test('scaled slide size changes all slides in one reversible history entry', async () => {
  const p = await fixture();
  recordEdit(p, { type: 'slide-duplicate', slide: 0 });
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const h = emptyHistory();
  h.entries.push(
    recordEdit(p, {
      type: 'slide-size',
      slide: 0,
      size: { width: before.width / 2, height: before.height / 2, type: 'custom' },
      scaleContent: true,
    }),
  );
  h.cursor = 1;
  const after = editorModel(p);
  for (const [index, slide] of after.slides.entries()) {
    const original = before.slides[index].shapes[0].bounds;
    assert.deepEqual(
      slide.shapes[0].bounds,
      Object.fromEntries(Object.entries(original).map(([key, value]) => [key, value / 2])),
    );
  }
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, h);
  assert.deepEqual(editorModel(replayed), after);
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(saved), after);
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...h, cursor: 0 });
  assert.deepEqual(editorModel(undone), before);
  const conflict = await pptx.loadPresentation(base);
  pptx.setShapeText(pptx.getSlideShapes(pptx.getSlides(conflict)[1])[0], 'Changed');
  assert.throws(() => replayEdits(conflict, h), /conflict/i);
  const changedMaster = unzipSync(base);
  const masterPath = Object.keys(changedMaster).find((name) =>
    /^ppt\/slideMasters\/slideMaster[^/]+\.xml$/.test(name),
  );
  const masterXml = strFromU8(changedMaster[masterPath]);
  const modifiedMaster = masterXml.replace(/sz="[0-9]+"/, 'sz="9900"');
  assert.notEqual(masterXml, modifiedMaster);
  changedMaster[masterPath] = strToU8(modifiedMaster);
  const masterConflict = await pptx.loadPresentation(zipSync(changedMaster));
  assert.throws(() => replayEdits(masterConflict, h), /conflict/i);
});

test('replace all preserves formatting, replays across slides and detects source conflicts', async () => {
  const p = await fixture();
  const first = pptx.getSlideShapes(pptx.getSlides(p)[0])[0];
  pptx.setShapeText(first, 'Alpha ALPHA');
  pptx.setShapeTextRangeFormat(first, 6, 11, { bold: true });
  pptx.addSlideTextBox(pptx.addBlankSlide(p), {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(1),
    text: 'alpha',
  });
  const base = await pptx.savePresentation(p);
  const history = emptyHistory();
  history.entries.push(
    recordEdit(p, {
      type: 'text-replace-all',
      slide: 0,
      replacement: { query: 'alpha', text: 'alphabet' },
    }),
  );
  history.cursor = 1;
  assert.deepEqual(
    editorModel(p).slides.map((s) => s.shapes[0].text),
    ['alphabet alphabet', 'alphabet'],
  );
  assert.ok(editorModel(p).slides[0].shapes[0].runs.some((r) => r.start === 9 && r.format.bold));
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, history);
  assert.deepEqual(editorModel(replay), editorModel(p));
  const conflict = await pptx.loadPresentation(base);
  pptx.setShapeText(pptx.getSlideShapes(pptx.getSlides(conflict)[1])[0], 'changed');
  assert.throws(() => replayEdits(conflict, history), /conflicts/);
  history.cursor = 0;
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, history);
  assert.deepEqual(
    editorModel(undo).slides.map((s) => s.shapes[0].text),
    ['Alpha ALPHA', 'alpha'],
  );
  assert.throws(
    () =>
      recordEdit(undo, {
        type: 'text-replace-all',
        slide: 0,
        replacement: { query: '', text: 'x' },
      }),
    /empty/,
  );
});

test('replace all applies search options during replay', async () => {
  const p = await fixture();
  const shape = pptx.getSlideShapes(pptx.getSlides(p)[0])[0];
  pptx.setShapeText(shape, 'Cat cat scatter');
  const base = await pptx.savePresentation(p);
  const record = recordEdit(p, {
    type: 'text-replace-all',
    slide: 0,
    replacement: {
      query: 'cat',
      text: 'dog',
      options: { matchCase: true, wholeWords: true },
    },
  });
  assert.equal(editorModel(p).slides[0].shapes[0].text, 'Cat dog scatter');
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, { ...emptyHistory(), entries: [record], cursor: 1 });
  assert.equal(editorModel(replay).slides[0].shapes[0].text, 'Cat dog scatter');
});

test('replace all validates cross-slide fit targets before mutation and persists measurements', async () => {
  const p = await fixture();
  const first = pptx.getSlideShapes(pptx.getSlides(p)[0])[0];
  pptx.setShapeTextAutoFit(first, 'normal');
  const second = pptx.addSlideTextBox(pptx.addBlankSlide(p), {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(1),
    text: 'Original',
  });
  pptx.setShapeTextAutoFit(second, 'shape');
  const base = await pptx.savePresentation(p);
  const model = editorModel(p);
  const fits = [
    {
      slide: 0,
      id: model.slides[0].shapes[0].id,
      autoFitParams: { fontScale: 0.5, lnSpcReduction: 0 },
    },
    {
      slide: 1,
      id: model.slides[1].shapes[0].id,
      bounds: { ...model.slides[1].shapes[0].bounds, h: 1828800 },
    },
  ];
  const command = {
    type: 'text-replace-all',
    slide: 0,
    replacement: { query: 'Original', text: 'Replacement' },
    textFits: fits,
  };
  assert.throws(
    () => recordEdit(p, { ...command, textFits: [...fits, { slide: 1, id: 9999 }] }),
    /target/,
  );
  assert.deepEqual(editorModel(p), model);
  const record = recordEdit(p, command);
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, { ...emptyHistory(), entries: [record], cursor: 1 });
  assert.deepEqual(editorModel(replay), editorModel(saved));
  assert.equal(editorModel(saved).slides[0].shapes[0].autoFitParams.fontScale, 0.5);
  assert.equal(editorModel(saved).slides[1].shapes[0].bounds.h, 1828800);
});

test('snapshot paste survives source deletion and replays independently through history', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const snapshot = Buffer.from(base).toString('base64');
  const id = editorModel(p).slides[0].shapes[0].id;
  const h = emptyHistory();
  h.entries.push(recordEdit(p, { type: 'delete', slide: 0, ids: [id] }));
  h.entries.push(recordEdit(p, { type: 'paste', slide: 0, source: 0, ids: [id], snapshot }));
  const pastedId = editorModel(p).slides[0].shapes[0].id;
  h.entries.push(
    recordEdit(p, { type: 'update', slide: 0, ids: [pastedId], changes: { text: 'Changed' } }),
  );
  h.entries.push(recordEdit(p, { type: 'paste', slide: 0, source: 0, ids: [id], snapshot }));
  h.cursor = h.entries.length;
  assert.deepEqual(
    editorModel(p).slides[0].shapes.map((shape) => shape.text),
    ['Changed', 'Original'],
  );
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, JSON.parse(JSON.stringify(h)));
  assert.deepEqual(
    editorModel(replayed).slides[0].shapes.map((shape) => shape.text),
    ['Changed', 'Original'],
  );
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...h, cursor: 1 });
  assert.equal(editorModel(undone).slides[0].shapes.length, 0);
  assert.throws(
    () => recordEdit(p, { type: 'paste', slide: 0, source: 0, ids: [id], snapshot, cut: true }),
    /cannot delete/,
  );
});

test('snapshot paste preserves stacking and grouped connector targets across slides', async () => {
  const p = await fixture();
  const edit = (command) => recordEdit(p, { slide: 0, ...command });
  const shapes = () => editorModel(p).slides[0].shapes;
  const targetId = shapes()[0].id;
  edit({ type: 'duplicate', ids: [targetId] });
  const siblingId = shapes()[1].id;
  edit({ type: 'group', ids: [targetId, siblingId] });
  const groupId = shapes()[0].id;
  edit({
    type: 'insert',
    preset: 'lineArrow',
    from: { x: 0, y: 0 },
    toPoint: { x: 6000000, y: 4000000 },
    connections: { start: { shapeId: targetId, siteIndex: 3 } },
  });
  const lineId = shapes()[1].id;
  const destination = pptx.addBlankSlide(p);
  pptx.addSlideTextBox(destination, {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(1),
    h: pptx.inches(1),
    text: 'Destination',
  });
  const base = await pptx.savePresentation(p);
  const snapshot = Buffer.from(base).toString('base64');
  const h = emptyHistory();
  h.entries.push(edit({ type: 'delete', ids: [groupId, lineId] }));
  h.entries.push(
    edit({
      type: 'paste',
      slide: 1,
      source: 0,
      ids: [lineId, groupId],
      snapshot,
    }),
  );
  h.cursor = h.entries.length;
  const pasted = editorModel(p).slides[1].shapes;
  assert.deepEqual(
    pasted.map((shape) => shape.kind),
    ['shape', 'group', 'connector'],
  );
  const all = pptx.getSlideShapes(destination);
  assert.equal(new Set(all.map(pptx.getShapeId)).size, all.length);
  const group = all.find((shape) => pptx.getShapeId(shape) === pasted[1].id);
  const childId = pptx.getShapeId(pptx.getGroupChildren(group)[0]);
  assert.deepEqual(pasted[2].connections.start, { shapeId: childId, siteIndex: 3 });
  assert.notEqual(childId, targetId);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, JSON.parse(JSON.stringify(h)));
  assert.deepEqual(editorModel(replayed), editorModel(p));
  const reloaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(reloaded), editorModel(p));
  // A connector pasted without its target must not attach to an unrelated shape
  // whose ID happens to match the original in the destination slide.
  edit({ type: 'paste', slide: 1, source: 0, ids: [lineId], snapshot });
  assert.equal(editorModel(p).slides[1].shapes.at(-1).connections.start, null);
});

test('table insertion validates dimensions and survives undo, replay and export', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const command = {
    type: 'insert',
    slide: 0,
    preset: 'table',
    table: { rows: 3, columns: 4 },
    changes: { bounds: { x: 914400, y: 1828800, w: 5486400, h: 1097280 } },
  };
  const before = editorModel(p);
  for (const table of [
    { rows: 0, columns: 2 },
    { rows: 2, columns: 76 },
    { rows: 1.5, columns: 2 },
    { rows: 2, columns: NaN },
  ]) {
    assert.throws(() => recordEdit(p, { ...command, table }), /Table dimensions/);
    assert.deepEqual(editorModel(p), before);
  }
  const record = recordEdit(p, command);
  const check = (p) => {
    const table = pptx.getSlideTables(pptx.getSlides(p)[0])[0];
    assert.equal(pptx.getTableRowHeights(table).length, 3);
    assert.equal(pptx.getTableColumnWidths(table).length, 4);
    assert.deepEqual(pptx.getShapeBounds(table), command.changes.bounds);
    assert.equal(pptx.getTableCellText(pptx.getTableCell(table, 2, 3)), '');
  };
  check(p);
  check(await pptx.loadPresentation(await pptx.savePresentation(p)));
  const history = { ...emptyHistory(), entries: [record], cursor: 1 };
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, history);
  check(replay);
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, { ...history, cursor: 0 });
  assert.equal(pptx.getSlideTables(pptx.getSlides(undo)[0]).length, 0);
});

test('table cell edits preserve rich text, validate atomically and replay merged anchors', async () => {
  const p = await fixture();
  const slide = pptx.getSlides(p)[0];
  const table = pptx.addSlideTable(slide, {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [
      ['Hello 😀', 'covered'],
      ['other', 'last'],
    ],
  });
  pptx.mergeTableCells(table, { row: 0, col: 0, rowSpan: 1, colSpan: 2 });
  const id = pptx.getShapeId(table);
  const model = () => editorModel(p).slides[0].shapes.find((shape) => shape.id === id).table;
  assert.equal(model().cells[0][0].span.gridSpan, 2);
  assert.equal(model().cells[0][1].span.hMerge, true);
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const edit = (cell) => recordEdit(p, { type: 'table-cell-text', slide: 0, ids: [id], cell });
  for (const cell of [
    { row: 0, column: 1, edits: [{ start: 0, end: 0, text: 'bad' }] },
    { row: -1, column: 0, edits: [] },
    { row: 9, column: 0, edits: [] },
    {
      row: 0,
      column: 0,
      edits: [
        { start: 0, end: 0, text: 'good' },
        { start: 100, end: 100, text: 'bad' },
      ],
    },
    { row: 0, column: 0, edits: [{ start: 7, end: 7, text: 'bad' }] },
  ]) {
    assert.throws(() => edit(cell));
    assert.deepEqual(editorModel(p), before);
  }
  const record = edit({
    row: 0,
    column: 0,
    edits: [
      { start: 0, end: 5, text: 'Welcome' },
      { start: 0, end: 0, text: '日本語\n' },
    ],
  });
  assert.equal(model().cells[0][0].text, '日本語\nWelcome 😀');
  assert.equal(model().cells[1][0].text, 'other');
  const history = { ...emptyHistory(), entries: [record], cursor: 1 };
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, JSON.parse(JSON.stringify(history)));
  assert.deepEqual(editorModel(replay), editorModel(p));
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(loaded), editorModel(p));
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undone), before);
});

test('cell character formatting and typed insertion formats replay through undo and save', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [['Hello', 'other']],
  });
  const id = pptx.getShapeId(table);
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const formatCommand = {
    type: 'table-cell-format',
    slide: 0,
    ids: [id],
    cell: { row: 0, column: 0, edits: [] },
    changes: { range: { start: 1, end: 4 }, format: { bold: true } },
  };
  assert.throws(() =>
    recordEdit(p, {
      ...formatCommand,
      changes: { ...formatCommand.changes, range: { start: 0, end: 100 } },
    }),
  );
  assert.deepEqual(editorModel(p), before);
  const formatted = recordEdit(p, formatCommand);
  const typed = recordEdit(p, {
    type: 'table-cell-text',
    slide: 0,
    ids: [id],
    cell: {
      row: 0,
      column: 0,
      edits: [{ start: 5, end: 5, text: '!', format: { italic: true } }],
    },
  });
  const check = (p) => {
    const cell = pptx.getTableCell(pptx.getSlideTables(pptx.getSlides(p)[0])[0], 0, 0);
    const runs = pptx.getTableCellParagraphs(cell)[0].elements;
    assert.deepEqual(
      runs.map((r) => r.text),
      ['H', 'ell', 'o', '!'],
    );
    assert.equal(runs[1].format.bold, true);
    assert.notEqual(runs[2].format.bold, true);
    assert.equal(runs[3].format.italic, true);
  };
  check(p);
  check(await pptx.loadPresentation(await pptx.savePresentation(p)));
  const history = { ...emptyHistory(), entries: [formatted, typed], cursor: 2 };
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, JSON.parse(JSON.stringify(history)));
  check(replay);
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undo), before);
});

test('cell paragraph alignment changes only touched paragraphs and replays through undo', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [['First\nSecond\nThird', 'other']],
  });
  const id = pptx.getShapeId(table);
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const command = {
    type: 'table-cell-align',
    slide: 0,
    ids: [id],
    cell: { row: 0, column: 0, edits: [] },
    changes: { range: { start: 6, end: 12 }, align: 'center' },
  };
  for (const changes of [
    { ...command.changes, align: 'invalid' },
    { ...command.changes, range: { start: 0, end: 100 } },
  ]) {
    assert.throws(() => recordEdit(p, { ...command, changes }));
    assert.deepEqual(editorModel(p), before);
  }
  const record = recordEdit(p, command);
  const check = (p) => {
    const table = pptx.getSlideTables(pptx.getSlides(p)[0])[0];
    const cell = pptx.getTableCell(table, 0, 0);
    assert.deepEqual(
      pptx.getTableCellParagraphs(cell).map((p) => p.align),
      [null, 'center', null],
    );
    assert.equal(pptx.getTableCellText(pptx.getTableCell(table, 0, 1)), 'other');
  };
  check(p);
  check(await pptx.loadPresentation(await pptx.savePresentation(p)));
  const history = { ...emptyHistory(), entries: [record], cursor: 1 };
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, JSON.parse(JSON.stringify(history)));
  check(replay);
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undo), before);
});

test('cell paragraph line spacing changes only touched paragraphs and replays through undo', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [['First\nSecond\nThird', 'other']],
  });
  const id = pptx.getShapeId(table);
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const command = {
    type: 'table-cell-line-spacing',
    slide: 0,
    ids: [id],
    cell: { row: 0, column: 0, edits: [] },
    changes: { range: { start: 6, end: 12 }, lineSpacing: 1.5 },
  };
  for (const changes of [
    { ...command.changes, lineSpacing: -1 },
    { ...command.changes, range: { start: 0, end: 100 } },
  ]) {
    assert.throws(() => recordEdit(p, { ...command, changes }));
    assert.deepEqual(editorModel(p), before);
  }
  const record = recordEdit(p, command);
  const check = (p) => {
    const table = pptx.getSlideTables(pptx.getSlides(p)[0])[0];
    const cell = pptx.getTableCell(table, 0, 0);
    assert.deepEqual(
      pptx.getTableCellParagraphs(cell).map((p) => p.properties.lineSpacing ?? null),
      [null, { kind: 'pct', value: 1.5 }, null],
    );
    assert.equal(pptx.getTableCellText(pptx.getTableCell(table, 0, 1)), 'other');
  };
  check(p);
  check(await pptx.loadPresentation(await pptx.savePresentation(p)));
  const history = { ...emptyHistory(), entries: [record], cursor: 1 };
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, JSON.parse(JSON.stringify(history)));
  check(replay);
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undo), before);
});

test('cell paragraph bullets change only touched paragraphs and replays through undo', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [['First\nSecond\nThird', 'other']],
  });
  const id = pptx.getShapeId(table);
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const command = {
    type: 'table-cell-bullets',
    slide: 0,
    ids: [id],
    cell: { row: 0, column: 0, edits: [] },
    changes: { range: { start: 6, end: 12 }, bullets: 'number' },
  };
  for (const changes of [
    { ...command.changes, bullets: 'invalid' },
    { ...command.changes, range: { start: 0, end: 100 } },
  ]) {
    assert.throws(() => recordEdit(p, { ...command, changes }));
    assert.deepEqual(editorModel(p), before);
  }
  const record = recordEdit(p, command);
  const check = (p) => {
    const table = pptx.getSlideTables(pptx.getSlides(p)[0])[0];
    const cell = pptx.getTableCell(table, 0, 0);
    assert.deepEqual(
      pptx.getTableCellParagraphs(cell).map((p) => p.properties.bullet ?? null),
      ['none', 'number', 'none'],
    );
    assert.equal(pptx.getTableCellText(pptx.getTableCell(table, 0, 1)), 'other');
  };
  check(p);
  check(await pptx.loadPresentation(await pptx.savePresentation(p)));
  const history = { ...emptyHistory(), entries: [record], cursor: 1 };
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, JSON.parse(JSON.stringify(history)));
  check(replay);
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undo), before);
});

test('appending a table row grows a rotated table without moving its top edge and replays', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(2),
    y: pptx.inches(2),
    w: pptx.inches(4),
    h: pptx.inches(3),
    rows: [
      ['First', 'Second'],
      ['Last', 'End'],
    ],
  });
  pptx.setTableRowHeight(table, 0, pptx.inches(1));
  pptx.setTableRowHeight(table, 1, pptx.inches(2));
  pptx.setShapeRotation(table, 90);
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const id = pptx.getShapeId(table);
  assert.throws(() =>
    recordEdit(p, { type: 'table-row-append', slide: 0, ids: [before.slides[0].shapes[0].id] }),
  );
  assert.deepEqual(editorModel(p), before);
  const record = recordEdit(p, { type: 'table-row-append', slide: 0, ids: [id] });
  const check = (p) => {
    const shape = editorModel(p).slides[0].shapes.find((s) => s.id === id);
    assert.deepEqual(
      shape.table.cells.map((row) => row.map((cell) => cell.text)),
      [
        ['First', 'Second'],
        ['Last', 'End'],
        ['', ''],
      ],
    );
    assert.deepEqual(shape.table.rowHeights, [914400, 1828800, 1828800]);
    assert.deepEqual(shape.bounds, { x: 914400, y: 914400, w: 3657600, h: 4572000 });
    assert.equal(shape.rotation, 90);
  };
  check(p);
  check(await pptx.loadPresentation(await pptx.savePresentation(p)));
  const history = { ...emptyHistory(), entries: [record], cursor: 1 };
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, JSON.parse(JSON.stringify(history)));
  check(replay);
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undo), before);
});

test('cell text direction and anchoring validate atomically, preserve neighbors and replay', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [['First\nSecond', 'Neighbor']],
  });
  const id = pptx.getShapeId(table);
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const command = {
    type: 'table-cell-layout',
    slide: 0,
    ids: [id],
    cell: { row: 0, column: 0, edits: [] },
    changes: { anchor: 'bottom', direction: 'vert270' },
  };
  for (const changes of [
    {},
    { anchor: 'center', direction: 'invalid' },
    { anchor: 'invalid', direction: 'vert' },
  ]) {
    assert.throws(() => recordEdit(p, { ...command, changes }));
    assert.deepEqual(editorModel(p), before);
  }
  const record = recordEdit(p, command);
  const check = (p) => {
    const result = editorModel(p).slides[0].shapes.find((s) => s.id === id);
    const original = before.slides[0].shapes.find((s) => s.id === id);
    assert.equal(result.table.cells[0][0].anchor, 'bottom');
    assert.equal(result.table.cells[0][0].direction, 'vert270');
    assert.equal(result.table.cells[0][0].text, 'First\nSecond');
    assert.deepEqual(result.table.cells[0][1], original.table.cells[0][1]);
    assert.deepEqual(result.bounds, original.bounds);
  };
  check(p);
  check(await pptx.loadPresentation(await pptx.savePresentation(p)));
  const history = { ...emptyHistory(), entries: [record], cursor: 1 };
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, JSON.parse(JSON.stringify(history)));
  check(replay);
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undo), before);
  for (const direction of ['vert', 'wordArtVert', 'horz']) {
    recordEdit(p, { ...command, changes: { direction, anchor: 'center' } });
    const result = editorModel(p).slides[0].shapes.find((s) => s.id === id).table.cells[0][0];
    assert.equal(result.direction, direction === 'horz' ? null : direction);
    assert.equal(result.anchor, 'center');
  }
});

test('table style options validate atomically, preserve cells and replay', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(1),
    y: pptx.inches(2),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [
      ['Header', 'Other'],
      ['Body', 'Value'],
    ],
  });
  const id = pptx.getShapeId(table);
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const command = { type: 'table-style', slide: 0, ids: [id] };
  for (const tableStyle of [
    undefined,
    null,
    [],
    {},
    { firstRow: false, bandRow: 'yes' },
    { unknown: true },
  ]) {
    assert.throws(() => recordEdit(p, { ...command, tableStyle }));
    assert.deepEqual(editorModel(p), before);
  }
  assert.throws(() =>
    recordEdit(p, {
      ...command,
      ids: [before.slides[0].shapes[0].id],
      tableStyle: { firstRow: true },
    }),
  );
  const tableStyle = {
    firstRow: false,
    lastRow: true,
    firstCol: true,
    lastCol: true,
    bandRow: false,
    bandCol: true,
  };
  const record = recordEdit(p, { ...command, tableStyle });
  const check = (p) => {
    const result = editorModel(p).slides[0].shapes.find((s) => s.id === id);
    const original = before.slides[0].shapes.find((s) => s.id === id);
    assert.deepEqual(result.table.style, tableStyle);
    assert.deepEqual(result.table.cells, original.table.cells);
    assert.deepEqual(result.bounds, original.bounds);
  };
  check(p);
  check(await pptx.loadPresentation(await pptx.savePresentation(p)));
  const history = { ...emptyHistory(), entries: [record], cursor: 1 };
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, JSON.parse(JSON.stringify(history)));
  check(replay);
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undo), before);
  recordEdit(p, { ...command, tableStyle: { bandRow: true } });
  assert.deepEqual(pptx.getTableStyleFlags(table), { ...tableStyle, bandRow: true });
});

test('table shading targets a cell or whole table, validates and persists explicit no fill', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(1),
    y: pptx.inches(2),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [
      ['Header', 'Other'],
      ['Body', 'Value'],
    ],
  });
  const id = pptx.getShapeId(table);
  const base = await pptx.savePresentation(p);
  const before = pptx.getSlideXmlString(pptx.getSlides(p)[0]);
  const command = {
    type: 'table-fill',
    slide: 0,
    ids: [id],
    cell: { row: 0, column: 0, edits: [] },
  };
  for (const tableFill of [undefined, 123, 'red', 'FFFFFFoops']) {
    assert.throws(() => recordEdit(p, { ...command, tableFill }));
    assert.equal(pptx.getSlideXmlString(pptx.getSlides(p)[0]), before);
  }
  assert.throws(() =>
    recordEdit(p, { ...command, tableFill: 'FF0000', cell: { row: -1, column: 0, edits: [] } }),
  );
  const record = recordEdit(p, { ...command, tableFill: 'FF0000' });
  assert.equal(pptx.getTableCellFill(pptx.getTableCell(table, 0, 0)), '#FF0000');
  assert.equal(pptx.getTableCellFill(pptx.getTableCell(table, 0, 1)), null);
  const clear = recordEdit(p, { ...command, tableFill: null });
  assert.equal(pptx.isTableCellNoFill(pptx.getTableCell(table, 0, 0)), true);
  const check = (p) => {
    const t = pptx.getSlideTables(pptx.getSlides(p)[0])[0];
    assert.equal(pptx.isTableCellNoFill(pptx.getTableCell(t, 0, 0)), true);
    assert.equal(pptx.isTableCellNoFill(pptx.getTableCell(t, 0, 1)), false);
    assert.equal(pptx.getTableCellText(pptx.getTableCell(t, 0, 0)), 'Header');
  };
  check(await pptx.loadPresentation(await pptx.savePresentation(p)));
  const history = { ...emptyHistory(), entries: [record, clear], cursor: 2 };
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, history);
  check(replay);
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, { ...history, cursor: 1 });
  assert.equal(
    pptx.getTableCellFill(pptx.getTableCell(pptx.getSlideTables(pptx.getSlides(undo)[0])[0], 0, 0)),
    '#FF0000',
  );
  recordEdit(p, { ...command, cell: undefined, tableFill: '00FF00' });
  for (const cell of pptx.getTableCells(table).flat()) {
    assert.equal(pptx.getTableCellFill(cell), '#00FF00');
    assert.equal(pptx.isTableCellNoFill(cell), false);
  }
});

test('table border commands synchronize shared edges and preserve merged perimeter segments', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(1),
    y: pptx.inches(2),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [
      ['Merged', 'A'],
      ['', 'B'],
    ],
  });
  pptx.mergeTableCells(table, { row: 0, col: 0, rowSpan: 2, colSpan: 1 });
  const id = pptx.getShapeId(table);
  const base = await pptx.savePresentation(p);
  const command = {
    type: 'table-borders',
    slide: 0,
    ids: [id],
    tableBorders: { mode: 'all', color: 'FF0000', weight: 3, dash: 'dash' },
  };
  const original = pptx.getSlideXmlString(pptx.getSlides(p)[0]);
  for (const patch of [
    { mode: 'invalid' },
    { color: 'bad' },
    { weight: NaN },
    { dash: 'invalid' },
  ]) {
    assert.throws(() =>
      recordEdit(p, { ...command, tableBorders: { ...command.tableBorders, ...patch } }),
    );
    assert.equal(pptx.getSlideXmlString(pptx.getSlides(p)[0]), original);
  }
  assert.throws(() => recordEdit(p, { ...command, cell: { row: 1, column: 0, edits: [] } }));
  const first = recordEdit(p, command);
  const border = (p, r, c, side) =>
    pptx.getTableCellBorders(
      p,
      pptx.getTableCell(pptx.getSlideTables(pptx.getSlides(p)[0])[0], r, c),
    )[side];
  assert.equal(border(p, 0, 0, 'right').color, '#FF0000');
  assert.equal(border(p, 1, 0, 'right').color, '#FF0000');
  const second = recordEdit(p, {
    ...command,
    cell: { row: 1, column: 1, edits: [] },
    tableBorders: { ...command.tableBorders, mode: 'left', color: '0000FF' },
  });
  assert.equal(border(p, 0, 0, 'right').color, '#FF0000');
  assert.equal(border(p, 1, 0, 'right').color, '#0000FF');
  assert.equal(border(p, 1, 1, 'left').color, '#0000FF');
  const segmentedSvg = renderSlideToSvg(p, pptx.getSlides(p)[0]);
  const blueLines = [...segmentedSvg.matchAll(/<line[^>]*stroke="#0000FF"[^>]*>/g)].map(
    (match) => match[0],
  );
  assert.equal(blueLines.length, 2);
  for (const line of blueLines) {
    assert.match(line, /x1="288(?:\.0+)?"/);
    assert.match(line, /y1="288(?:\.0+)?"/);
    assert.match(line, /y2="384(?:\.0+)?"/);
  }

  const history = { ...emptyHistory(), entries: [first, second], cursor: 2 };
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, history);
  assert.equal(border(replay, 1, 0, 'right').color, '#0000FF');
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, { ...history, cursor: 1 });
  assert.equal(border(undo, 1, 0, 'right').color, '#FF0000');
  recordEdit(p, { ...command, tableBorders: { ...command.tableBorders, mode: 'none' } });
  const reloaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  for (const [r, c, side] of [
    [0, 0, 'top'],
    [1, 0, 'bottom'],
    [1, 1, 'left'],
    [0, 1, 'right'],
  ])
    assert.equal(border(reloaded, r, c, side).noFill, true);
  const svg = renderSlideToSvg(reloaded, pptx.getSlides(reloaded)[0]);
  assert.equal(svg.includes('stroke="#FF0000"'), false);
  assert.equal(svg.includes('stroke="#9CA3AF"'), false);
});

test('table distribution preserves bounds, merged spans and exact grid totals through replay', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(1),
    y: pptx.inches(2),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [
      ['A', '', 'B'],
      ['C', 'D', 'E'],
    ],
  });
  pptx.setTableColumnWidth(table, 0, pptx.emu(100001));
  pptx.setTableColumnWidth(table, 1, pptx.emu(200000));
  pptx.setTableColumnWidth(table, 2, pptx.emu(600000));
  pptx.setTableRowHeight(table, 0, pptx.emu(300001));
  pptx.setTableRowHeight(table, 1, pptx.emu(700000));
  pptx.mergeTableCells(table, { row: 0, col: 0, rowSpan: 1, colSpan: 2 });
  pptx.setShapeRotation(table, 30);
  const box = pptx.getShapeBoundsResolved(p, table);
  const base = await pptx.savePresentation(p);
  const command = {
    type: 'table-distribute',
    slide: 0,
    ids: [pptx.getShapeId(table)],
    tableAxis: 'columns',
  };
  const xml = pptx.getSlideXmlString(pptx.getSlides(p)[0]);
  for (const patch of [
    { tableAxis: 'bad' },
    { cell: { row: 0, column: 1, edits: [] } },
    { cell: { row: -1, column: 0, edits: [] } },
  ]) {
    assert.throws(() => recordEdit(p, { ...command, ...patch }));
    assert.equal(pptx.getSlideXmlString(pptx.getSlides(p)[0]), xml);
  }
  const first = recordEdit(p, { ...command, cell: { row: 0, column: 0, edits: [] } });
  assert.deepEqual(pptx.getTableColumnWidths(table), [150001, 150000, 600000]);
  const second = recordEdit(p, { ...command, tableAxis: 'rows' });
  const check = (p) => {
    const table = pptx.getSlideTables(pptx.getSlides(p)[0])[0];
    assert.deepEqual(pptx.getTableColumnWidths(table), [150001, 150000, 600000]);
    assert.deepEqual(pptx.getTableRowHeights(table), [500001, 500000]);
    assert.deepEqual(pptx.getShapeBoundsResolved(p, table), box);
    assert.equal(pptx.getShapeRotation(table), 30);
    assert.equal(pptx.getTableCellSpan(pptx.getTableCell(table, 0, 0)).gridSpan, 2);
    assert.equal(pptx.getTableCellText(pptx.getTableCell(table, 1, 2)), 'E');
  };
  check(await pptx.loadPresentation(await pptx.savePresentation(p)));
  const history = { ...emptyHistory(), entries: [first, second], cursor: 2 };
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, history);
  check(replay);
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, { ...history, cursor: 1 });
  assert.deepEqual(
    pptx.getTableRowHeights(pptx.getSlideTables(pptx.getSlides(undo)[0])[0]),
    [300001, 700000],
  );
});

test('table cell sizing preserves neighbors and the rotated origin through serialization and undo', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(1),
    y: pptx.inches(2),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [
      ['A', '', 'B'],
      ['C', 'D', 'E'],
    ],
  });
  pptx.mergeTableCells(table, { row: 0, col: 0, rowSpan: 1, colSpan: 2 });
  pptx.setShapeRotation(table, 30);
  const base = await pptx.savePresentation(p);
  const old = pptx.getShapeBoundsResolved(p, table);
  const widths = pptx.getTableColumnWidths(table);
  const neighbor = (widths[2] * old.w) / widths.reduce((sum, v) => sum + v, 0);
  const command = {
    type: 'table-cell-size',
    slide: 0,
    ids: [pptx.getShapeId(table)],
    tableAxis: 'columns',
    tableCellSize: pptx.inches(2),
    cell: { row: 0, column: 0, edits: [] },
  };
  const xml = pptx.getSlideXmlString(pptx.getSlides(p)[0]);
  for (const size of [0, -1, NaN, Infinity, 51206401]) {
    assert.throws(() => recordEdit(p, { ...command, tableCellSize: size }));
    assert.equal(pptx.getSlideXmlString(pptx.getSlides(p)[0]), xml);
  }
  const entry = recordEdit(p, command);
  const check = (p) => {
    const table = pptx.getSlideTables(pptx.getSlides(p)[0])[0];
    const next = pptx.getTableColumnWidths(table);
    assert.equal(next[0] + next[1], pptx.inches(2));
    assert.equal(next[2], Math.round(neighbor));
    const box = pptx.getShapeBoundsResolved(p, table);
    const origin = (b) => ({
      x: b.x + b.w / 2 - (Math.cos(Math.PI / 6) * b.w) / 2 + (Math.sin(Math.PI / 6) * b.h) / 2,
      y: b.y + b.h / 2 - (Math.sin(Math.PI / 6) * b.w) / 2 - (Math.cos(Math.PI / 6) * b.h) / 2,
    });
    assert.ok(Math.abs(origin(box).x - origin(old).x) < 1);
    assert.ok(Math.abs(origin(box).y - origin(old).y) < 1);
    assert.equal(pptx.getTableCellText(pptx.getTableCell(table, 1, 2)), 'E');
  };
  check(await pptx.loadPresentation(await pptx.savePresentation(p)));
  const history = { ...emptyHistory(), entries: [entry], cursor: 1 };
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, history);
  check(replay);
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, { ...history, cursor: 0 });
  assert.deepEqual(
    pptx.getShapeBoundsResolved(undo, pptx.getSlideTables(pptx.getSlides(undo)[0])[0]),
    old,
  );
  recordEdit(p, {
    ...command,
    tableAxis: 'rows',
    cell: undefined,
    tableCellSize: pptx.inches(0.5),
  });
  assert.deepEqual(pptx.getTableRowHeights(table), [pptx.inches(0.5), pptx.inches(0.5)]);
  assert.equal(pptx.getShapeBoundsResolved(p, table).h, pptx.inches(1));
});

test('table cell margins validate atomically and preserve omitted sides through replay and undo', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(1),
    y: pptx.inches(2),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [
      ['A', 'B'],
      ['C', 'D'],
    ],
  });
  pptx.setTableCellMargins(pptx.getTableCell(table, 0, 0), { right: 12345 });
  const base = await pptx.savePresentation(p);
  const command = {
    type: 'table-margins',
    slide: 0,
    ids: [pptx.getShapeId(table)],
    cell: { row: 0, column: 0, edits: [] },
    tableMargins: { left: 182880, top: 0 },
  };
  const xml = pptx.getSlideXmlString(pptx.getSlides(p)[0]);
  for (const margins of [
    { left: 1, right: -1 },
    { left: NaN },
    { bottom: Infinity },
    { unknown: 1 },
    {},
  ]) {
    assert.throws(() => recordEdit(p, { ...command, tableMargins: margins }));
    assert.equal(pptx.getSlideXmlString(pptx.getSlides(p)[0]), xml);
  }
  const entry = recordEdit(p, command);
  const check = (p) => {
    const t = pptx.getSlideTables(pptx.getSlides(p)[0])[0];
    assert.deepEqual(pptx.getTableCellMargins(pptx.getTableCell(t, 0, 0)), {
      left: 182880,
      right: 12345,
      top: 0,
      bottom: null,
    });
    assert.equal(pptx.getTableCellMargins(pptx.getTableCell(t, 0, 1)).left, 91440);
    assert.equal(pptx.getTableCellText(pptx.getTableCell(t, 0, 0)), 'A');
  };
  check(await pptx.loadPresentation(await pptx.savePresentation(p)));
  const history = { ...emptyHistory(), entries: [entry], cursor: 1 };
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, history);
  check(replay);
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, { ...history, cursor: 0 });
  assert.equal(
    pptx.getTableCellMargins(
      pptx.getTableCell(pptx.getSlideTables(pptx.getSlides(undo)[0])[0], 0, 0),
    ).left,
    null,
  );
  recordEdit(p, { ...command, cell: undefined, tableMargins: { bottom: 45720 } });
  for (const cell of pptx.getTableCells(table).flat())
    assert.equal(pptx.getTableCellMargins(cell).bottom, 45720);
  assert.equal(pptx.getTableCellMargins(pptx.getTableCell(table, 0, 0)).right, 12345);
});

test('table insertion grows a rotated frame, preserves crossed merges and replays all four positions', async () => {
  for (const placement of ['above', 'below', 'left', 'right']) {
    const p = await fixture();
    const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
      x: pptx.inches(2),
      y: pptx.inches(2),
      w: pptx.inches(3),
      h: pptx.inches(3),
      rows: [
        ['A', 'B', 'C'],
        ['D', 'E', 'F'],
        ['G', 'H', 'I'],
      ],
    });
    pptx.setShapeRotation(table, 90);
    pptx.mergeTableCells(table, { row: 0, col: 0, rowSpan: 2, colSpan: 2 });
    const id = pptx.getShapeId(table);
    const base = await pptx.savePresentation(p);
    const before = editorModel(p);
    const rows = placement === 'above' || placement === 'below';
    // An unmerged cell beside the 2×2 merge inserts through that merge.
    const cell = rows
      ? { row: placement === 'above' ? 1 : 0, column: 2 }
      : { row: 2, column: placement === 'left' ? 1 : 0 };
    const command = {
      type: 'table-insert',
      slide: 0,
      ids: [id],
      tablePlacement: placement,
      cell: { ...cell, edits: [] },
    };
    assert.throws(() => recordEdit(p, { ...command, tablePlacement: 'invalid' }));
    assert.throws(() => recordEdit(p, { ...command, cell: { row: 1, column: 1, edits: [] } }));
    assert.deepEqual(editorModel(p), before);
    const record = recordEdit(p, command);
    const check = (p) => {
      const shape = editorModel(p).slides[0].shapes.find((s) => s.id === id);
      assert.equal(shape.table.rowHeights.length, rows ? 4 : 3);
      assert.equal(shape.table.columnWidths.length, rows ? 3 : 4);
      assert.equal(shape.table.cells[0][0].text, 'A');
      assert.equal(shape.table.cells[0][0].span.rowSpan, rows ? 3 : 2);
      assert.equal(shape.table.cells[0][0].span.gridSpan, rows ? 2 : 3);
      assert.deepEqual(shape.bounds, {
        x: 1371600,
        y: rows ? 1371600 : 2286000,
        w: rows ? 2743200 : 3657600,
        h: rows ? 3657600 : 2743200,
      });
      const inserted = rows ? shape.table.cells[1][2] : shape.table.cells[2][1];
      assert.equal(inserted.text, '');
    };
    check(p);
    check(await pptx.loadPresentation(await pptx.savePresentation(p)));
    const history = { ...emptyHistory(), entries: [record], cursor: 1 };
    const replay = await pptx.loadPresentation(base);
    replayEdits(replay, JSON.parse(JSON.stringify(history)));
    check(replay);
    const undo = await pptx.loadPresentation(base);
    replayEdits(undo, { ...history, cursor: 0 });
    assert.deepEqual(editorModel(undo), before);
  }
});

test('clearing a cell range retains the table and merges, rejects partial merges, and undoes', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(6),
    h: pptx.inches(2),
    rows: [
      ['A', 'B', 'C'],
      ['D', 'E', 'F'],
    ],
  });
  pptx.mergeTableCells(
    table,
    { row: 0, col: 0, rowSpan: 1, colSpan: 2 },
    { coveredText: 'append' },
  );
  pptx.setTableCellFill(getCell(1, 0), '#FF0000');
  function getCell(row, column) {
    return pptx.getTableCell(table, row, column);
  }
  const base = await pptx.savePresentation(p),
    before = editorModel(p);
  const shape = before.slides[0].shapes.find((s) => s.table);
  const command = {
    type: 'table-clear',
    slide: 0,
    ids: [shape.id],
    tableRange: { row: 0, column: 0, rows: 2, columns: 2 },
  };
  for (const range of [
    { row: 0, column: 1, rows: 2, columns: 1 },
    { row: 0, column: 0, rows: 2, columns: 4 },
    { row: 0, column: 0, rows: 0, columns: 1 },
  ]) {
    assert.throws(() => recordEdit(p, { ...command, tableRange: range }));
    assert.deepEqual(editorModel(p), before);
  }
  const history = emptyHistory();
  history.entries.push(recordEdit(p, command));
  history.cursor = 1;
  const after = editorModel(p),
    result = after.slides[0].shapes.find((s) => s.table);
  assert.deepEqual(
    result.table.cells.map((row) => row.map((cell) => cell.text)),
    [
      ['', '', 'C'],
      ['', '', 'F'],
    ],
  );
  assert.deepEqual(result.bounds, shape.bounds);
  assert.deepEqual(result.table.columnWidths, shape.table.columnWidths);
  assert.deepEqual(result.table.rowHeights, shape.table.rowHeights);
  assert.deepEqual(result.table.cells[0][0].span, shape.table.cells[0][0].span);
  assert.deepEqual(pptx.getTableCellFill(getCell(1, 0)), '#FF0000');
  assert.deepEqual(editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))), after);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed), after);
  history.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, history);
  assert.deepEqual(editorModel(undone), before);
});

test('table merge retains all cell text through reload, replay and undo', async () => {
  for (const premerged of [false, true]) {
    const p = await fixture();
    const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
      x: pptx.inches(1),
      y: pptx.inches(1),
      w: pptx.inches(6),
      h: pptx.inches(2),
      rows: [
        ['A', 'B', 'C'],
        ['D', 'E', 'F'],
      ],
    });
    if (premerged) {
      pptx.mergeTableCells(
        table,
        { row: 0, col: 0, rowSpan: 1, colSpan: 2 },
        { coveredText: 'append' },
      );
      pptx.mergeTableCells(
        table,
        { row: 1, col: 0, rowSpan: 1, colSpan: 2 },
        { coveredText: 'append' },
      );
    }
    const base = await pptx.savePresentation(p);
    const before = editorModel(p);
    const shape = before.slides[0].shapes.find((s) => s.table);
    const command = {
      type: 'table-merge',
      slide: 0,
      ids: [shape.id],
      tableRange: { row: 0, column: 0, rows: 2, columns: 2 },
    };
    assert.throws(() =>
      recordEdit(p, { ...command, tableRange: { ...command.tableRange, columns: 4 } }),
    );
    assert.deepEqual(editorModel(p), before);
    const history = emptyHistory();
    history.entries.push(recordEdit(p, command));
    history.cursor = 1;
    const after = editorModel(p);
    const result = after.slides[0].shapes.find((s) => s.table);
    assert.equal(result.table.cells[0][0].text, 'A\nB\nD\nE');
    assert.equal(result.table.cells[1][1].text, '');
    assert.equal(result.table.cells[0][0].span.gridSpan, 2);
    assert.equal(result.table.cells[0][0].span.rowSpan, 2);
    assert.deepEqual(result.bounds, shape.bounds);
    assert.deepEqual(
      editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))),
      after,
    );
    const replayed = await pptx.loadPresentation(base);
    replayEdits(replayed, history);
    assert.deepEqual(editorModel(replayed), after);
    history.cursor = 0;
    const undone = await pptx.loadPresentation(base);
    replayEdits(undone, history);
    assert.deepEqual(editorModel(undone), before);
  }
});

test('table deletion shrinks rotated grids, repairs crossing merges, reloads, replays and undoes', async () => {
  for (const tableAxis of ['rows', 'columns']) {
    const p = await fixture();
    const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
      x: pptx.inches(1),
      y: pptx.inches(1),
      w: pptx.inches(6),
      h: pptx.inches(3),
      rows: [
        ['A', 'B', 'C'],
        ['D', 'E', 'F'],
        ['G', 'H', 'I'],
      ],
    });
    pptx.mergeTableCells(table, { row: 0, col: 0, rowSpan: 2, colSpan: 2 });
    pptx.setShapeRotation(table, 90);
    const base = await pptx.savePresentation(p);
    const before = editorModel(p);
    const shape = before.slides[0].shapes.find((s) => s.table);
    const cell =
      tableAxis === 'rows' ? { row: 0, column: 2, edits: [] } : { row: 2, column: 0, edits: [] };
    const command = { type: 'table-delete', slide: 0, ids: [shape.id], tableAxis, cell };
    assert.throws(() => recordEdit(p, { ...command, tableAxis: 'invalid' }));
    assert.throws(() => recordEdit(p, { ...command, cell: { row: 1, column: 1, edits: [] } }));
    assert.deepEqual(editorModel(p), before);
    const history = emptyHistory();
    history.entries.push(recordEdit(p, command));
    history.cursor = 1;
    const after = editorModel(p);
    const result = after.slides[0].shapes.find((s) => s.table);
    assert.equal(result.table.cells.length, tableAxis === 'rows' ? 2 : 3);
    assert.equal(result.table.cells[0].length, tableAxis === 'rows' ? 3 : 2);
    assert.equal(result.table.cells[0][0].text, 'A');
    assert.equal(result.table.cells[0][0].span.rowSpan, tableAxis === 'rows' ? 1 : 2);
    assert.equal(result.table.cells[0][0].span.gridSpan, tableAxis === 'rows' ? 2 : 1);
    assert.equal(result.bounds.w, shape.bounds.w * (tableAxis === 'columns' ? 2 / 3 : 1));
    assert.equal(result.bounds.h, shape.bounds.h * (tableAxis === 'rows' ? 2 / 3 : 1));
    assert.equal(
      result.bounds.x + (result.bounds.w + result.bounds.h) / 2,
      shape.bounds.x + (shape.bounds.w + shape.bounds.h) / 2,
    );
    assert.equal(
      result.bounds.y + (result.bounds.h - result.bounds.w) / 2,
      shape.bounds.y + (shape.bounds.h - shape.bounds.w) / 2,
    );
    const merged = await pptx.loadPresentation(base);
    recordEdit(merged, { ...command, cell: { row: 0, column: 0, edits: [] } });
    const remaining = editorModel(merged).slides[0].shapes.find((s) => s.table).table;
    assert.equal(remaining.cells.length, tableAxis === 'rows' ? 1 : 3);
    assert.equal(remaining.columnWidths.length, tableAxis === 'columns' ? 1 : 3);
    assert.equal(remaining.cells[0][0].text, tableAxis === 'rows' ? 'G' : 'C');
    assert.deepEqual(
      editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))),
      after,
    );
    const replay = await pptx.loadPresentation(base);
    replayEdits(replay, history);
    assert.deepEqual(editorModel(replay), after);
    history.cursor = 0;
    const undo = await pptx.loadPresentation(base);
    replayEdits(undo, history);
    assert.deepEqual(editorModel(undo), before);
    recordEdit(p, { ...command, cell: undefined });
    assert.equal(
      editorModel(p).slides[0].shapes.some((s) => s.table),
      false,
    );
  }
});

test('cell range formatting preserves neighbors and supports merged and empty cells', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(6),
    h: pptx.inches(2),
    rows: [
      ['A', 'B', 'C'],
      ['', 'E', 'F'],
    ],
  });
  pptx.mergeTableCells(
    table,
    { row: 0, col: 0, rowSpan: 1, colSpan: 2 },
    { coveredText: 'append' },
  );
  const base = await pptx.savePresentation(p),
    before = editorModel(p);
  const id = before.slides[0].shapes.find((s) => s.table).id;
  const range = { row: 0, column: 0, rows: 2, columns: 2 };
  const changes = {
    format: { bold: true, color: '#FF0000', size: 24 },
    align: 'center',
    bullets: 'bullet',
    lineSpacing: 1.5,
    anchor: 'bottom',
    direction: 'vert',
  };
  for (const invalid of [
    { tableRange: { ...range, column: 1 } },
    { changes: { ...changes, anchor: 'invalid' } },
    { changes: { ...changes, lineSpacing: -1 } },
  ]) {
    assert.throws(() =>
      recordEdit(p, {
        type: 'table-range-format',
        slide: 0,
        ids: [id],
        tableRange: range,
        changes,
        ...invalid,
      }),
    );
    assert.deepEqual(editorModel(p), before);
  }
  const history = emptyHistory();
  for (const edit of [
    { type: 'table-range-format', changes },
    { type: 'table-fill', tableFill: '#00FF00' },
    { type: 'table-margins', tableMargins: { left: 12345 } },
  ])
    history.entries.push(recordEdit(p, { ...edit, slide: 0, ids: [id], tableRange: range }));
  history.cursor = history.entries.length;
  const after = editorModel(p),
    model = after.slides[0].shapes.find((s) => s.table);
  for (const [row, column] of [
    [0, 0],
    [1, 0],
    [1, 1],
  ]) {
    const cell = pptx.getTableCell(table, row, column),
      data = model.table.cells[row][column];
    assert.equal(pptx.getTableCellFill(cell), '#00FF00');
    assert.equal(data.margins.left, 12345);
    assert.equal(data.anchor, 'bottom');
    assert.equal(data.direction, 'vert');
    for (const para of data.paragraphs) {
      assert.equal(para.align, 'center');
      const formats = para.elements.filter((e) => e.kind !== 'br').map((e) => e.format);
      if (!formats.length) formats.push(para.endFormat);
      assert.ok(formats.length);
      for (const format of formats) {
        assert.equal(format.bold, true);
        assert.equal(format.size, 24);
      }
    }
  }
  for (const row of [0, 1])
    assert.deepEqual(
      model.table.cells[row][2],
      before.slides[0].shapes.find((s) => s.table).table.cells[row][2],
    );
  assert.deepEqual(editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))), after);
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, history);
  assert.deepEqual(editorModel(replay), after);
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undo), before);
});

test('range borders target rectangle edges, synchronize neighbors, and survive replay', async () => {
  for (const mode of [
    'all',
    'none',
    'outside',
    'inside',
    'horizontal',
    'vertical',
    'left',
    'right',
    'top',
    'bottom',
    'tlToBr',
    'blToTr',
  ]) {
    const p = await fixture();
    const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
      x: pptx.inches(1),
      y: pptx.inches(1),
      w: pptx.inches(6),
      h: pptx.inches(3),
      rows: Array.from({ length: 4 }, () => ['A', 'B', 'C', 'D']),
    });
    const id = pptx.getShapeId(table);
    recordEdit(p, {
      type: 'table-borders',
      slide: 0,
      ids: [id],
      tableBorders: { mode: 'all', color: '0000FF', weight: 1, dash: 'solid' },
    });
    const base = await pptx.savePresentation(p),
      before = editorModel(p);
    const command = {
      type: 'table-borders',
      slide: 0,
      ids: [id],
      tableRange: { row: 1, column: 1, rows: 2, columns: 2 },
      tableBorders: { mode, color: 'FF0000', weight: 3, dash: 'dash' },
    };
    const record = recordEdit(p, command);
    const border = (r, c, side) =>
      pptx.getTableCellBorders(p, pptx.getTableCell(table, r, c))[side];
    const check = (value, changed) => {
      if (changed && mode === 'none') assert.equal(value.noFill, true);
      else assert.equal(value.color, changed ? '#FF0000' : '#0000FF', mode);
    };
    for (let r = 0; r < 4; r++)
      for (let c = 0; c <= 4; c++) {
        const inRange = r >= 1 && r < 3 && c >= 1 && c <= 3;
        const changed =
          inRange &&
          (['all', 'none'].includes(mode) ||
            (mode === 'outside' && c !== 2) ||
            (['inside', 'vertical'].includes(mode) && c === 2) ||
            (mode === 'left' && c === 1) ||
            (mode === 'right' && c === 3));
        if (c > 0) check(border(r, c - 1, 'right'), changed);
        if (c < 4) check(border(r, c, 'left'), changed);
      }
    for (let r = 0; r <= 4; r++)
      for (let c = 0; c < 4; c++) {
        const inRange = c >= 1 && c < 3 && r >= 1 && r <= 3;
        const changed =
          inRange &&
          (['all', 'none'].includes(mode) ||
            (mode === 'outside' && r !== 2) ||
            (['inside', 'horizontal'].includes(mode) && r === 2) ||
            (mode === 'top' && r === 1) ||
            (mode === 'bottom' && r === 3));
        if (r > 0) check(border(r - 1, c, 'bottom'), changed);
        if (r < 4) check(border(r, c, 'top'), changed);
      }
    if (mode === 'tlToBr' || mode === 'blToTr') {
      for (let r = 0; r < 4; r++)
        for (let c = 0; c < 4; c++)
          assert.equal(border(r, c, mode)?.color === '#FF0000', r >= 1 && r < 3 && c >= 1 && c < 3);
    }
    const after = editorModel(p),
      history = { ...emptyHistory(), entries: [record], cursor: 1 };
    assert.deepEqual(
      editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))),
      after,
    );
    const replay = await pptx.loadPresentation(base);
    replayEdits(replay, history);
    assert.deepEqual(editorModel(replay), after);
    const undo = await pptx.loadPresentation(base);
    replayEdits(undo, { ...history, cursor: 0 });
    assert.deepEqual(editorModel(undo), before);
  }
});

test('range borders reject partial merged selections before changing any edge', async () => {
  const p = await fixture(),
    table = pptx.addSlideTable(pptx.getSlides(p)[0], {
      x: pptx.inches(1),
      y: pptx.inches(1),
      w: pptx.inches(4),
      h: pptx.inches(2),
      rows: [
        ['A', 'B', 'C'],
        ['D', 'E', 'F'],
      ],
    });
  pptx.mergeTableCells(table, { row: 0, col: 0, rowSpan: 1, colSpan: 2 });
  const before = editorModel(p),
    command = {
      type: 'table-borders',
      slide: 0,
      ids: [pptx.getShapeId(table)],
      tableBorders: { mode: 'outside', color: 'FF0000', weight: 1, dash: 'solid' },
    };
  for (const range of [
    { row: 0, column: 0, rows: 2, columns: 1 },
    { row: 0, column: 1, rows: 2, columns: 2 },
    { row: 0, column: 0, rows: 3, columns: 3 },
    { row: 0, column: 0, rows: 0, columns: 2 },
  ]) {
    assert.throws(() => recordEdit(p, { ...command, tableRange: range }));
    assert.deepEqual(editorModel(p), before);
  }
  recordEdit(p, { ...command, tableRange: { row: 0, column: 0, rows: 2, columns: 2 } });
  assert.equal(pptx.getTableCellBorders(p, pptx.getTableCell(table, 0, 1)).right.color, '#FF0000');
  assert.equal(pptx.getTableCellBorders(p, pptx.getTableCell(table, 0, 2)).left.color, '#FF0000');
});

test('range sizing and distribution preserve unselected grid sizes and rotated origin', async () => {
  for (const axis of ['rows', 'columns'])
    for (const type of ['table-cell-size', 'table-distribute']) {
      const p = await fixture(),
        table = pptx.addSlideTable(pptx.getSlides(p)[0], {
          x: pptx.inches(1),
          y: pptx.inches(1),
          w: pptx.inches(10),
          h: pptx.inches(10),
          rows: Array.from({ length: 4 }, () => ['A', 'B', 'C', 'D']),
        });
      pptx.setShapeRotation(table, 30);
      const set = axis === 'rows' ? pptx.setTableRowHeight : pptx.setTableColumnWidth;
      const get = axis === 'rows' ? pptx.getTableRowHeights : pptx.getTableColumnWidths;
      [1, 2, 3, 4].forEach((value, index) => set(table, index, pptx.inches(value)));
      const base = await pptx.savePresentation(p),
        before = editorModel(p),
        sizes = get(table);
      const command = {
        type,
        slide: 0,
        ids: [pptx.getShapeId(table)],
        tableAxis: axis,
        tableCellSize: pptx.inches(1.5),
        tableRange: { row: 1, column: 1, rows: 2, columns: 2 },
      };
      assert.throws(() =>
        recordEdit(p, { ...command, tableRange: { ...command.tableRange, rows: 5 } }),
      );
      assert.deepEqual(editorModel(p), before);
      const record = recordEdit(p, command),
        after = editorModel(p),
        result = get(table);
      assert.equal(result[0], sizes[0]);
      assert.equal(result[3], sizes[3]);
      assert.equal(result[1], pptx.inches(type === 'table-distribute' ? 2.5 : 1.5));
      assert.equal(result[2], result[1]);
      const old = before.slides[0].shapes.find((s) => s.table).bounds,
        next = after.slides[0].shapes.find((s) => s.table).bounds;
      const corner = (box) => ({
        x:
          box.x +
          box.w / 2 -
          (box.w / 2) * Math.cos(Math.PI / 6) +
          (box.h / 2) * Math.sin(Math.PI / 6),
        y:
          box.y +
          box.h / 2 -
          (box.w / 2) * Math.sin(Math.PI / 6) -
          (box.h / 2) * Math.cos(Math.PI / 6),
      });
      assert.ok(Math.abs(corner(old).x - corner(next).x) < 1);
      assert.ok(Math.abs(corner(old).y - corner(next).y) < 1);
      assert.deepEqual(
        editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))),
        after,
      );
      const replay = await pptx.loadPresentation(base);
      replayEdits(replay, { ...emptyHistory(), entries: [record], cursor: 1 });
      assert.deepEqual(editorModel(replay), after);
      const undo = await pptx.loadPresentation(base);
      replayEdits(undo, { ...emptyHistory(), entries: [record], cursor: 0 });
      assert.deepEqual(editorModel(undo), before);
    }
});

test('sizing a selected merged cell uses its full width and rejects partial merges', async () => {
  const p = await fixture(),
    table = pptx.addSlideTable(pptx.getSlides(p)[0], {
      x: pptx.inches(1),
      y: pptx.inches(1),
      w: pptx.inches(6),
      h: pptx.inches(2),
      rows: [
        ['A', 'B', 'C'],
        ['D', 'E', 'F'],
      ],
    });
  pptx.mergeTableCells(table, { row: 0, col: 0, rowSpan: 1, colSpan: 2 });
  const before = editorModel(p),
    command = {
      type: 'table-cell-size',
      slide: 0,
      ids: [pptx.getShapeId(table)],
      tableAxis: 'columns',
      tableCellSize: pptx.inches(3),
      tableRange: { row: 0, column: 0, rows: 1, columns: 2 },
    };
  assert.throws(() =>
    recordEdit(p, { ...command, tableRange: { ...command.tableRange, columns: 1 } }),
  );
  assert.deepEqual(editorModel(p), before);
  recordEdit(p, command);
  assert.deepEqual(pptx.getTableColumnWidths(table), [
    pptx.inches(1.5),
    pptx.inches(1.5),
    pptx.inches(2),
  ]);
  assert.equal(pptx.getTableCellSpan(pptx.getTableCell(table, 0, 0)).gridSpan, 2);
});

test('table split fits content and preserves rotated top edge through reload, replay and undo', async () => {
  for (const merged of [false, true]) {
    const p = await fixture();
    const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
      rows: [
        ['A', 'B'],
        ['C', 'D'],
      ],
      x: pptx.inches(1),
      y: pptx.inches(1),
      w: pptx.inches(6),
      h: pptx.inches(2),
    });
    pptx.setShapeRotation(table, 30);
    if (merged)
      pptx.mergeTableCells(
        table,
        { row: 0, col: 0, rowSpan: 2, colSpan: 2 },
        { coveredText: 'append' },
      );
    const before = editorModel(p),
      base = await pptx.savePresentation(p);
    const shape = before.slides[0].shapes.find((s) => s.table);
    const command = {
      type: 'table-split',
      slide: 0,
      ids: [shape.id],
      cell: { row: 0, column: 0, edits: [] },
      tableSplit: { rows: 2, columns: 3 },
    };
    assert.throws(() => recordEdit(p, { ...command, tableSplit: { rows: 0, columns: 3 } }));
    assert.deepEqual(editorModel(p), before);
    const history = emptyHistory();
    history.entries.push(recordEdit(p, command));
    history.cursor = 1;
    const after = editorModel(p),
      result = after.slides[0].shapes.find((s) => s.table);
    if (!merged) assert.deepEqual(result.bounds, shape.bounds);
    else {
      assert.ok(result.bounds.h > shape.bounds.h);
      const growth = result.bounds.h - shape.bounds.h;
      assert.ok(
        Math.abs(result.bounds.x - shape.bounds.x + (Math.sin(Math.PI / 6) * growth) / 2) <= 1,
      );
      assert.ok(
        Math.abs(result.bounds.y - shape.bounds.y - ((Math.cos(Math.PI / 6) - 1) * growth) / 2) <=
          1,
      );
      assert.equal(result.bounds.w, shape.bounds.w);
    }
    assert.equal(result.rotation, shape.rotation);
    assert.equal(result.table.cells[0][0].text, merged ? 'A\nB\nC\nD' : 'A');
    assert.deepEqual(
      editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))),
      after,
    );
    const replayed = await pptx.loadPresentation(base);
    replayEdits(replayed, history);
    assert.deepEqual(editorModel(replayed), after);
    history.cursor = 0;
    const undone = await pptx.loadPresentation(base);
    replayEdits(undone, history);
    assert.deepEqual(editorModel(undone), before);
  }
});

test('table minimum height accounts for wrapping, margins and empty paragraph font size', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: 0,
    y: 0,
    w: pptx.inches(4),
    h: pptx.inches(1),
    rows: [['A', '']],
  });
  const cell = pptx.getTableCell(table, 0, 0);
  const short = measureTableCellHeight(p, table, 0, 0);
  pptx.setTableCellText(
    cell,
    'Many words that must wrap across several lines inside this table cell.',
  );
  const wrapped = measureTableCellHeight(p, table, 0, 0);
  assert.ok(wrapped > short * 2);
  pptx.setTableCellMargins(cell, { top: pptx.inches(0.5), bottom: pptx.inches(0.5) });
  assert.ok(measureTableCellHeight(p, table, 0, 0) > wrapped + pptx.inches(0.8));
  const empty = pptx.getTableCell(table, 0, 1);
  const small = measureTableCellHeight(p, table, 0, 1);
  pptx.setTableCellTextFormat(empty, { size: 60 });
  assert.ok(measureTableCellHeight(p, table, 0, 1) > small * 2);
  pptx.setTableCellTextDirection(cell, 'vert');
  assert.equal(measureTableCellHeight(p, table, 0, 0), null);
});

test('table text growth fits a merged row span once and is stable on replay and undo', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(1),
    rows: [
      ['A', 'B'],
      ['', 'C'],
    ],
  });
  pptx.mergeTableCells(table, { row: 0, col: 0, rowSpan: 2, colSpan: 1 });
  const before = editorModel(p),
    base = await pptx.savePresentation(p);
  const entry = recordEdit(p, {
    type: 'table-cell-text',
    slide: 0,
    ids: [pptx.getShapeId(table)],
    cell: { row: 0, column: 0, edits: [{ start: 0, end: 1, text: 'A\nB\nC\nD\nE\nF' }] },
  });
  const height = pptx.getTableRowHeights(table).reduce((a, b) => a + b, 0);
  const minimum = measureTableCellHeight(p, table, 0, 0);
  assert.ok(height >= minimum && height <= minimum + 2);
  assert.equal(measureTableCellHeight(p, table, 1, 0), null);
  const after = editorModel(p);
  const history = { ...emptyHistory(), entries: [entry], cursor: 1 };
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, history);
  assert.deepEqual(editorModel(replay), after);
  assert.deepEqual(editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))), after);
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undo), before);
});

test('table distribution respects content minimums while keeping selected rows equal', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(3),
    rows: [
      ['outside', ''],
      ['A\nB\nC\nD\nE\nF', ''],
      ['short', ''],
      ['outside', ''],
    ],
  });
  pptx.setShapeRotation(table, 30);
  const old = pptx.getShapeBoundsResolved(p, table);
  const initial = pptx.getTableRowHeights(table);
  const before = editorModel(p),
    base = await pptx.savePresentation(p);
  const entry = recordEdit(p, {
    type: 'table-distribute',
    slide: 0,
    ids: [pptx.getShapeId(table)],
    tableAxis: 'rows',
    tableRange: { row: 1, column: 0, rows: 2, columns: 2 },
  });
  const rows = pptx.getTableRowHeights(table);
  assert.equal(rows[0], initial[0]);
  assert.equal(rows[3], initial[3]);
  assert.ok(rows[1] > initial[1]);
  assert.ok(Math.abs(rows[1] - rows[2]) <= 1);
  assert.ok(rows[1] >= measureTableCellHeight(p, table, 1, 0));
  const next = pptx.getShapeBoundsResolved(p, table),
    growth = next.h - old.h;
  assert.ok(Math.abs(next.x - old.x + (Math.sin(Math.PI / 6) * growth) / 2) <= 1);
  assert.ok(Math.abs(next.y - old.y - ((Math.cos(Math.PI / 6) - 1) * growth) / 2) <= 1);
  const after = editorModel(p),
    history = { ...emptyHistory(), entries: [entry], cursor: 1 };
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, history);
  assert.deepEqual(editorModel(replay), after);
  assert.deepEqual(editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))), after);
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undo), before);
});

test('table column distribution and merged-row deletion retain sufficient text height', async () => {
  for (const operation of ['columns', 'delete']) {
    const p = await fixture();
    const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
      x: 0,
      y: 0,
      w: pptx.inches(4),
      h: pptx.inches(1),
      rows: [
        ['Words that wrap when this cell becomes narrower than its original width', ''],
        ['', ''],
      ],
    });
    pptx.mergeTableCells(table, { row: 0, col: 0, rowSpan: 2, colSpan: 1 });
    pptx.setTableColumnWidth(table, 0, pptx.inches(3.5));
    pptx.setTableColumnWidth(table, 1, pptx.inches(0.5));
    recordEdit(p, {
      type: operation === 'columns' ? 'table-distribute' : 'table-delete',
      slide: 0,
      ids: [pptx.getShapeId(table)],
      tableAxis: operation === 'columns' ? 'columns' : 'rows',
      ...(operation === 'delete' ? { tableRange: { row: 1, column: 0, rows: 1, columns: 2 } } : {}),
    });
    const total = pptx.getTableRowHeights(table).reduce((a, b) => a + b, 0);
    assert.ok(total >= measureTableCellHeight(p, table, 0, 0));
    if (operation === 'columns')
      assert.deepEqual(pptx.getTableColumnWidths(table), [pptx.inches(2), pptx.inches(2)]);
    else assert.equal(pptx.getTableRowHeights(table).length, 1);
  }
});

test('resizing tables fits wrapped text and retains the opposite rotated edge', async () => {
  for (const anchor of ['top', 'bottom']) {
    const p = await fixture();
    const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
      x: pptx.inches(1),
      y: pptx.inches(1),
      w: pptx.inches(6),
      h: pptx.inches(2),
      rows: [['Several words that need multiple lines when resized narrower', 'neighbor']],
    });
    pptx.setShapeRotation(table, 30);
    const id = pptx.getShapeId(table),
      before = editorModel(p),
      base = await pptx.savePresentation(p);
    const requested = {
      x: pptx.inches(2),
      y: pptx.inches(2),
      w: pptx.inches(2),
      h: pptx.inches(0.2),
    };
    const command = {
      type: 'update',
      slide: 0,
      ids: [id],
      tableResizeAnchor: anchor,
      ...(anchor === 'top'
        ? { changes: { bounds: requested } }
        : { positions: [{ id, bounds: requested }] }),
    };
    assert.throws(() => recordEdit(p, { ...command, tableResizeAnchor: 'invalid' }));
    assert.deepEqual(editorModel(p), before);
    const entry = recordEdit(p, command);
    const actual = pptx.getShapeBoundsResolved(p, table);
    assert.ok(actual.h >= measureTableCellHeight(p, table, 0, 0));
    assert.ok(actual.h > requested.h);
    assert.equal(actual.w, requested.w);
    const direction = anchor === 'top' ? 1 : -1,
      growth = actual.h - requested.h;
    assert.ok(
      Math.abs(actual.x - requested.x + (Math.sin(Math.PI / 6) * direction * growth) / 2) <= 1,
    );
    assert.ok(
      Math.abs(actual.y - requested.y - ((Math.cos(Math.PI / 6) * direction - 1) * growth) / 2) <=
        1,
    );
    const after = editorModel(p),
      history = { ...emptyHistory(), entries: [entry], cursor: 1 };
    const replay = await pptx.loadPresentation(base);
    replayEdits(replay, history);
    assert.deepEqual(editorModel(replay), after);
    assert.deepEqual(
      editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))),
      after,
    );
    const undo = await pptx.loadPresentation(base);
    replayEdits(undo, { ...history, cursor: 0 });
    assert.deepEqual(editorModel(undo), before);
  }
});

test('moving or rotating a table leaves authored row heights unchanged', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: 0,
    y: 0,
    w: pptx.inches(3),
    h: pptx.inches(0.2),
    rows: [['A\nB\nC', '']],
  });
  const heights = pptx.getTableRowHeights(table),
    box = pptx.getShapeBoundsResolved(p, table);
  recordEdit(p, {
    type: 'update',
    slide: 0,
    ids: [pptx.getShapeId(table)],
    positions: [
      { id: pptx.getShapeId(table), bounds: { ...box, x: pptx.inches(1), y: pptx.inches(2) } },
    ],
    changes: { rotation: 30 },
  });
  assert.deepEqual(pptx.getTableRowHeights(table), heights);
  assert.equal(pptx.getShapeBoundsResolved(p, table).h, box.h);
});

test('table range insertion adds selected counts at each range edge and survives replay', async () => {
  for (const placement of ['above', 'below', 'left', 'right']) {
    const p = await fixture();
    const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
      x: pptx.inches(1),
      y: pptx.inches(1),
      w: pptx.inches(4),
      h: pptx.inches(4),
      rows: Array.from({ length: 4 }, (_, r) => Array.from({ length: 4 }, (_, c) => `${r},${c}`)),
    });
    const id = pptx.getShapeId(table);
    const base = await pptx.savePresentation(p);
    const before = editorModel(p);
    const command = {
      type: 'table-insert',
      slide: 0,
      ids: [id],
      tablePlacement: placement,
      tableRange: { row: 1, column: 1, rows: 2, columns: 2 },
    };
    for (const invalid of [
      { row: 3, column: 1, rows: 2, columns: 2 },
      { row: 1, column: 1, rows: 0, columns: 2 },
    ]) {
      assert.throws(() => recordEdit(p, { ...command, tableRange: invalid }));
      assert.deepEqual(editorModel(p), before);
    }
    const record = recordEdit(p, command);
    const rows = placement === 'above' || placement === 'below';
    const index = placement === 'below' || placement === 'right' ? 3 : 1;
    const check = (deck) => {
      const shape = editorModel(deck).slides[0].shapes.find((s) => s.id === id);
      assert.equal(shape.table.rowHeights.length, rows ? 6 : 4);
      assert.equal(shape.table.columnWidths.length, rows ? 4 : 6);
      for (let r = 0; r < 4; r++)
        for (let c = 0; c < 4; c++) {
          assert.equal(
            shape.table.cells[r + (rows && r >= index ? 2 : 0)][c + (!rows && c >= index ? 2 : 0)]
              .text,
            `${r},${c}`,
          );
        }
      for (let offset = 0; offset < 2; offset++)
        assert.equal(
          shape.table.cells[rows ? index + offset : 0][rows ? 0 : index + offset].text,
          '',
        );
    };
    check(p);
    check(await pptx.loadPresentation(await pptx.savePresentation(p)));
    const replay = await pptx.loadPresentation(base);
    const history = { ...emptyHistory(), entries: [record], cursor: 1 };
    replayEdits(replay, JSON.parse(JSON.stringify(history)));
    check(replay);
    const undo = await pptx.loadPresentation(base);
    replayEdits(undo, { ...history, cursor: 0 });
    assert.deepEqual(editorModel(undo), before);
  }
});

test('paragraph dialog settings target selected shape paragraphs and replay through undo', async () => {
  const p = await fixture();
  const shape = pptx.getSlideShapes(pptx.getSlides(p)[0])[0];
  pptx.setShapeParagraphs(shape, [
    { align: 'center', runs: [{ text: 'First', format: { bold: true } }] },
    { align: 'right', runs: [{ text: 'Second', format: { italic: true } }] },
  ]);
  const before = editorModel(p),
    bytes = await pptx.savePresentation(p),
    history = emptyHistory();
  const id = before.slides[0].shapes[0].id;
  history.entries.push(
    recordEdit(p, {
      type: 'update',
      slide: 0,
      ids: [id],
      changes: {
        range: { start: 6, end: 12 },
        paragraph: {
          leftEmu: 182880,
          firstLineEmu: -91440,
          beforePts: 6,
          afterPts: 8,
          lineSpacing: { kind: 'pct', value: 1.5 },
        },
      },
    }),
  );
  history.cursor++;
  const after = editorModel(p);
  assert.deepEqual(
    after.slides[0].shapes[0].paragraphs[0],
    before.slides[0].shapes[0].paragraphs[0],
  );
  const properties = after.slides[0].shapes[0].paragraphs[1].properties;
  assert.equal(properties.marL, 182880);
  assert.equal(properties.indent, -91440);
  assert.equal(properties.spcBefPts, 6);
  assert.equal(properties.spcAftPts, 8);
  assert.deepEqual(properties.lineSpacing, { kind: 'pct', value: 1.5 });
  assert.deepEqual(after.slides[0].shapes[0].runs, before.slides[0].shapes[0].runs);
  const replayed = await pptx.loadPresentation(bytes);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed), after);
  history.cursor = 0;
  const undone = await pptx.loadPresentation(bytes);
  replayEdits(undone, history);
  assert.deepEqual(editorModel(undone), before);
});

test('selected case conversion retains mixed runs across Unicode expansion', async () => {
  const p = await fixture();
  const shape = pptx.getSlideShapes(pptx.getSlides(p)[0])[0];
  pptx.setShapeText(shape, 'aßz tail');
  pptx.setShapeTextRangeFormat(shape, 1, 2, { bold: true, size: 30 });
  const id = pptx.getShapeId(shape);
  recordEdit(p, {
    type: 'update',
    slide: 0,
    ids: [id],
    changes: { changeCase: 'upper', range: { start: 0, end: 3 } },
  });
  const result = editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))).slides[0]
    .shapes[0];
  assert.equal(result.text, 'ASSZ tail');
  const expanded = result.runs.filter((run) => run.start < 3 && run.end > 1);
  assert.ok(expanded.length);
  assert.ok(expanded.every((run) => run.format.bold && run.format.size === 30));
  assert.notEqual(result.runs.at(-1).format.size, 30);
});

test('outline edits retain color during width changes and support export, replay and undo', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const id = before.slides[0].shapes[0].id;
  const history = emptyHistory();
  for (const changes of [
    { stroke: { color: '#13579B', widthEmu: 12700 } },
    { stroke: { widthEmu: 38100 }, strokeDash: 'lgDashDot', strokeOpacity: 0.35 },
    { stroke: { color: '#13579B' }, strokeCap: 'sq', strokeJoin: 'bevel', strokeCompound: 'dbl' },
  ])
    history.entries.push(recordEdit(p, { type: 'update', slide: 0, ids: [id], changes }));
  history.cursor = 3;
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(saved).slides[0].shapes[0].paint.stroke, {
    kind: 'solid',
    color: '#13579B',
    widthEmu: 38100,
  });
  assert.equal(editorModel(saved).slides[0].shapes[0].paint.dash, 'lgDashDot');
  assert.equal(editorModel(saved).slides[0].shapes[0].paint.strokeOpacity, 0.35);
  const paint = editorModel(saved).slides[0].shapes[0].paint;
  assert.equal(paint.cap, 'sq');
  assert.equal(paint.join, 'bevel');
  assert.equal(paint.compound, 'dbl');
  const svg = renderSlideToSvg(saved, pptx.getSlides(saved)[0]);
  assert.match(svg, /stroke-linecap="square"/);
  assert.match(svg, /stroke-linejoin="bevel"/);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed), editorModel(saved));
  history.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, history);
  assert.deepEqual(editorModel(undone), before);
});

test('arrow edits preserve per-object sizes and opposite ends through export and undo', async () => {
  const p = await fixture();
  const slide = pptx.getSlides(p)[0];
  const lines = [0, 1].map((i) =>
    pptx.addSlideLine(slide, {
      from: { x: pptx.inches(1), y: pptx.inches(1 + i) },
      to: { x: pptx.inches(3), y: pptx.inches(1 + i) },
    }),
  );
  for (const [i, line] of lines.entries()) {
    pptx.setShapeStroke(line, { color: '#13579B', widthEmu: 38100 });
    pptx.setShapeStrokeArrow(line, 'head', {
      type: 'triangle',
      width: i ? 'lg' : 'sm',
      length: 'lg',
    });
    pptx.setShapeStrokeArrow(line, 'tail', { type: 'oval', width: 'med', length: 'sm' });
  }
  const ids = lines.map(pptx.getShapeId);
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const history = emptyHistory();
  for (const changes of [
    { strokeHeadArrow: { type: 'stealth' } },
    { strokeTailArrow: { width: 'lg', length: 'med' } },
  ])
    history.entries.push(recordEdit(p, { type: 'update', slide: 0, ids, changes }));
  history.cursor = history.entries.length;
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  const shapes = editorModel(saved).slides[0].shapes.filter((shape) => ids.includes(shape.id));
  for (const [i, shape] of shapes.entries()) {
    assert.deepEqual(shape.paint.headArrow, {
      type: 'stealth',
      width: i ? 'lg' : 'sm',
      length: 'lg',
    });
    assert.deepEqual(shape.paint.tailArrow, { type: 'oval', width: 'lg', length: 'med' });
    assert.deepEqual(shape.paint.stroke, { kind: 'solid', color: '#13579B', widthEmu: 38100 });
  }
  const svg = renderSlideToSvg(saved, pptx.getSlides(saved)[0]);
  assert.match(svg, /marker-start="url\(#/);
  assert.match(svg, /marker-end="url\(#/);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed), editorModel(saved));
  history.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, history);
  assert.deepEqual(editorModel(undone), before);
});

test('picture crop shape preserves image and geometry through export, replay and undo', async () => {
  const p = await fixture();
  const picture = pptx.addSlideImage(
    pptx.getSlides(p)[0],
    Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jp1sAAAAASUVORK5CYII=',
      'base64',
    ),
    { x: 100000, y: 200000, w: 2000000, h: 1000000 },
  );
  pptx.setShapeImageCrop(picture, { top: 0.2, left: 0.1 });
  pptx.setShapeRotation(picture, 30);
  pptx.setShapeStroke(picture, { color: '#FF0000', widthEmu: 38100 });
  const id = pptx.getShapeId(picture);
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const history = emptyHistory();
  for (const preset of ['ellipse', 'roundRect'])
    history.entries.push(
      recordEdit(p, { type: 'update', slide: 0, ids: [id], changes: { imageCropShape: preset } }),
    );
  history.cursor = history.entries.length;
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  const after = editorModel(saved);
  assert.equal(after.slides[0].shapes.at(-1).preset, 'roundRect');
  assert.deepEqual(
    { ...after.slides[0].shapes.at(-1), preset: 'rect' },
    before.slides[0].shapes.at(-1),
  );
  const savedShape = pptx
    .getSlideShapes(pptx.getSlides(saved)[0])
    .find((s) => pptx.getShapeId(s) === id);
  assert.deepEqual(pptx.getShapeImageCrop(savedShape), pptx.getShapeImageCrop(picture));
  assert.deepEqual(
    Array.from(pptx.getShapeImageBytes(savedShape)),
    Array.from(pptx.getShapeImageBytes(picture)),
  );
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed), after);
  history.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, history);
  assert.deepEqual(editorModel(undone), before);
  assert.throws(() => pptx.setShapeImageCropShape(savedShape, 'bogus'), /preset/);
  assert.throws(
    () => pptx.setShapeImageCropShape(pptx.getSlideShapes(pptx.getSlides(saved)[0])[0], 'ellipse'),
    /picture/,
  );
});

test('picture Fill and Fit survive export, replay and undo with an unchanged frame', async () => {
  const p = await fixture();
  const picture = pptx.addSlideImage(
    pptx.getSlides(p)[0],
    Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jp1sAAAAASUVORK5CYII=',
      'base64',
    ),
    { x: 100000, y: 200000, w: 2000000, h: 1000000 },
  );
  const id = pptx.getShapeId(picture),
    base = await pptx.savePresentation(p);
  const before = editorModel(p),
    history = emptyHistory();
  for (const imageFit of ['fill', 'fit']) {
    history.entries.push(
      recordEdit(p, { type: 'update', slide: 0, ids: [id], changes: { imageFit } }),
    );
    history.cursor++;
    const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
    const after = editorModel(saved);
    assert.deepEqual(
      after.slides[0].shapes.at(-1).imageCrop,
      imageFit === 'fill'
        ? { left: 0, right: 0, top: 0.25, bottom: 0.25 }
        : { left: -0.5, right: -0.5, top: 0, bottom: 0 },
    );
    assert.deepEqual(after.slides[0].shapes.at(-1).bounds, before.slides[0].shapes.at(-1).bounds);
    const replayed = await pptx.loadPresentation(base);
    replayEdits(replayed, history);
    assert.deepEqual(editorModel(replayed), after);
  }
  history.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, history);
  assert.deepEqual(editorModel(undone), before);
});

test('picture crop aspect ratio survives export, replay and undo', async () => {
  const p = await fixture();
  const picture = pptx.addSlideImage(
    pptx.getSlides(p)[0],
    Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jp1sAAAAASUVORK5CYII=',
      'base64',
    ),
    { x: 100000, y: 200000, w: 2000000, h: 1000000 },
  );
  const id = pptx.getShapeId(picture),
    base = await pptx.savePresentation(p);
  const before = editorModel(p),
    history = emptyHistory();
  for (const imageCropAspectRatio of [1, 2 / 3, 16 / 9]) {
    history.entries.push(
      recordEdit(p, { type: 'update', slide: 0, ids: [id], changes: { imageCropAspectRatio } }),
    );
    history.cursor++;
    const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
    const after = editorModel(saved);
    const bounds = after.slides[0].shapes.at(-1).bounds;
    assert.ok(Math.abs(bounds.w / bounds.h - imageCropAspectRatio) < 0.00001);
    const replayed = await pptx.loadPresentation(base);
    replayEdits(replayed, history);
    assert.deepEqual(editorModel(replayed), after);
  }
  history.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, history);
  assert.deepEqual(editorModel(undone), before);
});

test('picture transparency survives export, replay and undo', async () => {
  const p = await fixture();
  const picture = pptx.addSlideImage(
    pptx.getSlides(p)[0],
    Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jp1sAAAAASUVORK5CYII=',
      'base64',
    ),
    { x: 0, y: 0, w: 1000000, h: 1000000 },
  );
  const id = pptx.getShapeId(picture),
    base = await pptx.savePresentation(p),
    before = editorModel(p),
    history = emptyHistory();
  for (const imageOpacity of [0, 0.73, 1]) {
    history.entries.push(
      recordEdit(p, { type: 'update', slide: 0, ids: [id], changes: { imageOpacity } }),
    );
    history.cursor++;
    const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
    assert.equal(editorModel(saved).slides[0].shapes.at(-1).imageOpacity, imageOpacity);
    const replayed = await pptx.loadPresentation(base);
    replayEdits(replayed, history);
    assert.deepEqual(editorModel(replayed), editorModel(saved));
  }
  const conflict = await pptx.loadPresentation(base);
  pptx.setShapeImageOpacity(pptx.getSlideShapes(pptx.getSlides(conflict)[0]).at(-1), 0.2);
  assert.throws(() => replayEdits(conflict, history), /conflict/i);
  const unchanged = editorModel(p);
  for (const imageOpacity of [-1, 2, NaN, Infinity]) {
    assert.throws(() =>
      recordEdit(p, { type: 'update', slide: 0, ids: [id], changes: { imageOpacity } }),
    );
    assert.deepEqual(editorModel(p), unchanged);
  }
  history.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, history);
  assert.deepEqual(editorModel(undone), before);
});

for (const [key, setter] of [
  ['imageBrightness', pptx.setShapeImageBrightness],
  ['imageContrast', pptx.setShapeImageContrast],
]) {
  test(`picture ${key} survives export, replay and undo`, async () => {
    const p = await fixture();
    const picture = pptx.addSlideImage(
      pptx.getSlides(p)[0],
      Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jp1sAAAAASUVORK5CYII=',
        'base64',
      ),
      { x: 0, y: 0, w: 1000000, h: 1000000 },
    );
    const id = pptx.getShapeId(picture),
      base = await pptx.savePresentation(p),
      before = editorModel(p),
      history = emptyHistory();
    for (const imageOpacity of [-1, 0.73, 1]) {
      history.entries.push(
        recordEdit(p, { type: 'update', slide: 0, ids: [id], changes: { [key]: imageOpacity } }),
      );
      history.cursor++;
      const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
      assert.equal(editorModel(saved).slides[0].shapes.at(-1)[key], imageOpacity);
      const replayed = await pptx.loadPresentation(base);
      replayEdits(replayed, history);
      assert.deepEqual(editorModel(replayed), editorModel(saved));
    }
    const conflict = await pptx.loadPresentation(base);
    setter(pptx.getSlideShapes(pptx.getSlides(conflict)[0]).at(-1), 0.2);
    assert.throws(() => replayEdits(conflict, history), /conflict/i);
    const unchanged = editorModel(p);
    for (const imageOpacity of [-2, 2, NaN, Infinity]) {
      assert.throws(() =>
        recordEdit(p, { type: 'update', slide: 0, ids: [id], changes: { [key]: imageOpacity } }),
      );
      assert.deepEqual(editorModel(p), unchanged);
    }
    history.cursor = 0;
    const undone = await pptx.loadPresentation(base);
    replayEdits(undone, history);
    assert.deepEqual(editorModel(undone), before);
  });
}

test('picture replacement isolates shared images and supports replay, conflict detection and undo', async () => {
  const p = await fixture();
  const original = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jp1sAAAAASUVORK5CYII=',
    'base64',
  );
  const replacement = Buffer.from(
    'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7',
    'base64',
  );
  const slide = pptx.getSlides(p)[0];
  const picture = pptx.addSlideImage(slide, original, {
    x: 10000,
    y: 20000,
    w: 1000000,
    h: 700000,
  });
  pptx.setShapeImageCrop(picture, { left: 0.1, right: 0.2, top: 0.05, bottom: 0 });
  pptx.setShapeImageOpacity(picture, 0.73);
  pptx.setShapeImageBrightness(picture, 0.2);
  pptx.setShapeRotation(picture, 30);
  const peer = pptx.copyShape(slide, picture);
  pptx.copyShape(pptx.addBlankSlide(p), picture);
  const id = pptx.getShapeId(picture);
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const history = emptyHistory();
  const command = {
    type: 'picture-replace',
    slide: 0,
    ids: [id],
    image: { base64: replacement.toString('base64'), name: 'replacement.gif' },
  };
  history.entries.push(recordEdit(p, command));
  history.cursor++;
  assert.deepEqual(
    editorModel(p).slides.map((slide) => slide.shapes),
    before.slides.map((slide) => slide.shapes),
  );
  const verify = (document, expected) => {
    const slides = pptx.getSlides(document);
    const pictures = pptx
      .getSlideShapes(slides[0])
      .filter((shape) => pptx.getShapeKind(shape) === 'picture');
    assert.deepEqual(Buffer.from(pptx.getShapeImageBytes(pictures[0])), expected);
    assert.deepEqual(Buffer.from(pptx.getShapeImageBytes(pictures[1])), original);
    assert.deepEqual(
      Buffer.from(pptx.getShapeImageBytes(pptx.getSlideShapes(slides[1])[0])),
      original,
    );
  };
  verify(await pptx.loadPresentation(await pptx.savePresentation(p)), replacement);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  verify(replayed, replacement);
  const conflict = await pptx.loadPresentation(base);
  pptx.setShapeImage(pptx.getSlideShapes(pptx.getSlides(conflict)[0])[1], replacement, {
    isolated: true,
  });
  assert.throws(() => replayEdits(conflict, history), /conflict/i);
  for (const base64 of ['', '!', 'AAAA']) {
    assert.throws(() => recordEdit(p, { ...command, image: { base64, name: 'bad' } }));
    verify(p, replacement);
  }
  // Exercise same-format isolation as well as cross-format replacement.
  recordEdit(p, {
    ...command,
    ids: [pptx.getShapeId(peer)],
    image: { base64: replacement.toString('base64'), name: 'peer.gif' },
  });
  recordEdit(p, {
    ...command,
    image: {
      base64: Buffer.from(replacement.map((b, i) => (i === 13 ? 255 : b))).toString('base64'),
      name: 'other.gif',
    },
  });
  assert.deepEqual(
    Buffer.from(
      pptx.getShapeImageBytes(
        pptx
          .getSlideShapes(pptx.getSlides(p)[0])
          .find((shape) => pptx.getShapeId(shape) === pptx.getShapeId(peer)),
      ),
    ),
    replacement,
  );
  history.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, history);
  verify(undone, original);
});

test('picture replacement isolates a shared relationship for same-format images', async () => {
  const p = await fixture();
  const original = Buffer.from(
    'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7',
    'base64',
  );
  const replacement = Buffer.from(original);
  replacement[13] = 255;
  const slide = pptx.getSlides(p)[0];
  const first = pptx.addSlideImage(slide, original, { x: 0, y: 0, w: 1000000, h: 1000000 });
  pptx.copyShape(slide, first);
  const zip = unzipSync(await pptx.savePresentation(p));
  let embed;
  zip['ppt/slides/slide1.xml'] = strToU8(
    strFromU8(zip['ppt/slides/slide1.xml']).replace(/r:embed="([^"]+)"/g, (_match, id) => {
      embed ??= id;
      return `r:embed="${embed}"`;
    }),
  );
  const shared = await pptx.loadPresentation(zipSync(zip));
  const pictures = pptx
    .getSlideShapes(pptx.getSlides(shared)[0])
    .filter((shape) => pptx.getShapeKind(shape) === 'picture');
  recordEdit(shared, {
    type: 'picture-replace',
    slide: 0,
    ids: [pptx.getShapeId(pictures[0])],
    image: { base64: replacement.toString('base64'), name: 'new.gif' },
  });
  const saved = await pptx.loadPresentation(await pptx.savePresentation(shared));
  const result = pptx
    .getSlideShapes(pptx.getSlides(saved)[0])
    .filter((shape) => pptx.getShapeKind(shape) === 'picture');
  assert.deepEqual(Buffer.from(pptx.getShapeImageBytes(result[0])), replacement);
  assert.deepEqual(Buffer.from(pptx.getShapeImageBytes(result[1])), original);
});

test('object links persist, replay, undo and detect source link conflicts', async () => {
  const p = await fixture();
  pptx.addBlankSlide(p);
  const base = await pptx.savePresentation(p);
  const model = editorModel(p);
  const id = model.slides[0].shapes[0].id;
  const history = emptyHistory();
  const url = {
    action: { kind: 'url', url: 'https://example.com/?a=1&b=2' },
    tooltip: '資料 <詳細>',
  };
  const destination = {
    action: { kind: 'slide', slide: model.slides[1].key },
    tooltip: '次の資料',
  };
  for (const link of [url, destination, { action: { kind: 'lastSlide' }, tooltip: null }, null]) {
    history.entries.push(recordEdit(p, { type: 'object-link', slide: 0, ids: [id], link }));
    history.cursor++;
    assert.deepEqual(editorModel(p).slides[0].shapes[0].link, link);
    const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
    assert.deepEqual(editorModel(saved).slides[0].shapes[0].link, link);
    const replayed = await pptx.loadPresentation(base);
    replayEdits(replayed, history);
    assert.deepEqual(editorModel(replayed).slides[0].shapes[0].link, link);
  }
  history.cursor = 2;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, history);
  assert.deepEqual(editorModel(undone).slides[0].shapes[0].link, destination);
  history.cursor = 0;
  const original = await pptx.loadPresentation(base);
  replayEdits(original, history);
  assert.equal(editorModel(original).slides[0].shapes[0].link, null);
  history.cursor = 1;
  const conflict = await pptx.loadPresentation(base);
  pptx.setShapeClickAction(pptx.getSlideShapes(pptx.getSlides(conflict)[0])[0], {
    kind: 'nextSlide',
  });
  assert.throws(() => replayEdits(conflict, history), /conflict/i);
});

test('object link validation is atomic and keeps text-run links separate', async () => {
  const p = await fixture();
  const shape = pptx.getSlideShapes(pptx.getSlides(p)[0])[0];
  pptx.setShapeHyperlink(shape, 'https://text.example.com/', 'Text tip');
  const id = pptx.getShapeId(shape);
  const command = { type: 'object-link', slide: 0, ids: [id] };
  const link = {
    action: { kind: 'url', url: 'https://object.example.com/' },
    tooltip: 'Object tip',
  };
  recordEdit(p, { ...command, link });
  assert.deepEqual(editorModel(p).slides[0].shapes[0].link, link);
  for (const invalid of [
    undefined,
    {},
    { action: { kind: 'url', url: '' }, tooltip: null },
    { action: { kind: 'slide', slide: '/ppt/slides/missing.xml' }, tooltip: null },
    { action: { kind: 'unsupported' }, tooltip: null },
    { action: { kind: 'nextSlide' }, tooltip: 42 },
  ]) {
    assert.throws(() => recordEdit(p, { ...command, link: invalid }));
    assert.deepEqual(editorModel(p).slides[0].shapes[0].link, link);
  }
  assert.throws(() => recordEdit(p, { ...command, ids: [], link }));
  assert.throws(() => recordEdit(p, { ...command, ids: [id, 999999], link: null }));
  recordEdit(p, { ...command, link: null });
  const restored = await pptx.loadPresentation(await pptx.savePresentation(p));
  const restoredShape = pptx.getSlideShapes(pptx.getSlides(restored)[0])[0];
  assert.equal(editorModel(restored).slides[0].shapes[0].link, null);
  assert.equal(pptx.getShapeHyperlink(restoredShape), 'https://text.example.com/');
  assert.equal(pptx.getShapeHyperlinkTooltip(restoredShape), 'Text tip');
});

test('text links retain ranges through save, replay, undo and conflicts', async () => {
  const p = await fixture();
  const shape = pptx.getSlideShapes(pptx.getSlides(p)[0])[0];
  pptx.setShapeText(shape, 'A😀B\n日本語');
  const id = pptx.getShapeId(shape);
  const base = await pptx.savePresentation(p);
  const link = { action: { kind: 'url', url: 'https://example.com/text' }, tooltip: '詳細' };
  const command = { type: 'text-link', slide: 0, ids: [id], range: { start: 1, end: 7 }, link };
  const history = emptyHistory();
  history.entries.push(recordEdit(p, command));
  history.cursor = 1;
  const expected = [
    { start: 1, end: 4, link },
    { start: 5, end: 7, link },
  ];
  assert.deepEqual(editorModel(p).slides[0].shapes[0].textLinks, expected);
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(saved).slides[0].shapes[0].textLinks, expected);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed).slides[0].shapes[0].textLinks, expected);
  const conflict = await pptx.loadPresentation(base);
  pptx.setShapeTextRangeClickAction(pptx.getSlideShapes(pptx.getSlides(conflict)[0])[0], 1, 3, {
    kind: 'nextSlide',
  });
  assert.throws(() => replayEdits(conflict, history), /conflict/i);
  history.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, history);
  assert.deepEqual(editorModel(undone).slides[0].shapes[0].textLinks, []);
  for (const range of [
    undefined,
    { start: 2, end: 3 },
    { start: 0, end: 99 },
    { start: 3, end: 1 },
    { start: 1, end: 1 },
  ]) {
    assert.throws(() => recordEdit(p, { ...command, range }));
    assert.deepEqual(editorModel(p).slides[0].shapes[0].textLinks, expected);
  }
  recordEdit(p, { ...command, range: { start: 1, end: 3 }, link: null });
  assert.deepEqual(editorModel(p).slides[0].shapes[0].textLinks, [
    { start: 3, end: 4, link },
    { start: 5, end: 7, link },
  ]);
});

test('cell text links persist and undo replacement, isolate cells and reject invalid targets', async () => {
  const p = await fixture();
  const table = pptx.addSlideTable(pptx.getSlides(p)[0], {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(4),
    h: pptx.inches(2),
    rows: [['A😀B', 'neighbor']],
  });
  const id = pptx.getShapeId(table);
  const base = await pptx.savePresentation(p);
  const link = { action: { kind: 'url', url: 'https://example.com/cell' }, tooltip: 'Cell' };
  const command = {
    type: 'text-link',
    slide: 0,
    ids: [id],
    cell: { row: 0, column: 0, edits: [] },
    range: { start: 1, end: 3 },
    text: '参照',
    link,
  };
  const history = emptyHistory();
  history.entries.push(recordEdit(p, command));
  history.cursor = 1;
  const cells = (p) => editorModel(p).slides[0].shapes.find((shape) => shape.id === id).table.cells;
  assert.equal(cells(p)[0][0].text, 'A参照B');
  assert.deepEqual(cells(p)[0][0].textLinks, [{ start: 1, end: 3, link }]);
  assert.deepEqual(cells(p)[0][1].textLinks, []);
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(cells(saved), cells(p));
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(cells(replayed), cells(p));
  const conflict = await pptx.loadPresentation(base);
  pptx.setTableCellTextRangeClickAction(
    pptx.getTableCell(pptx.getSlideTables(pptx.getSlides(conflict)[0])[0], 0, 0),
    1,
    3,
    { kind: 'nextSlide' },
  );
  assert.throws(() => replayEdits(conflict, history), /conflict/i);
  history.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, history);
  assert.equal(cells(undone)[0][0].text, 'A😀B');
  assert.deepEqual(cells(undone)[0][0].textLinks, []);
  const before = pptx.getSlideXmlString(pptx.getSlides(p)[0]);
  for (const cell of [
    undefined,
    { row: -1, column: 0 },
    { row: 0.5, column: 0 },
    { row: 5, column: 0 },
  ]) {
    assert.throws(() => recordEdit(p, { ...command, cell }));
    assert.equal(pptx.getSlideXmlString(pptx.getSlides(p)[0]), before);
  }
  recordEdit(p, { ...command, text: undefined, link: null });
  assert.equal(cells(p)[0][0].text, 'A参照B');
  assert.deepEqual(cells(p)[0][0].textLinks, []);
});

test('click and hover actions save atomically, replay, undo and detect hover conflicts', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const id = editorModel(p).slides[0].shapes[0].id;
  const link = { action: { kind: 'nextSlide' }, tooltip: 'Click' };
  const hoverLink = { action: { kind: 'endShow' }, tooltip: 'Hover' };
  const command = { type: 'object-actions', slide: 0, ids: [id], link, hoverLink };
  const history = emptyHistory();
  history.entries.push(recordEdit(p, command));
  history.cursor = 1;
  for (const target of [p, await pptx.loadPresentation(await pptx.savePresentation(p))]) {
    const shape = editorModel(target).slides[0].shapes[0];
    assert.deepEqual(shape.link, link);
    assert.deepEqual(shape.hoverLink, hoverLink);
  }
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed).slides[0].shapes[0].hoverLink, hoverLink);
  history.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, history);
  assert.equal(editorModel(undone).slides[0].shapes[0].link, null);
  assert.equal(editorModel(undone).slides[0].shapes[0].hoverLink, null);
  for (const invalid of [
    undefined,
    { action: { kind: 'invalid' }, tooltip: null },
    { action: { kind: 'slide', slide: '/missing' }, tooltip: null },
  ]) {
    const before = pptx.getSlideXmlString(pptx.getSlides(p)[0]);
    assert.throws(() => recordEdit(p, { ...command, link: null, hoverLink: invalid }));
    assert.equal(pptx.getSlideXmlString(pptx.getSlides(p)[0]), before);
  }
  history.cursor = 1;
  const conflict = await pptx.loadPresentation(base);
  pptx.setShapeHoverAction(pptx.getSlideShapes(pptx.getSlides(conflict)[0])[0], {
    kind: 'lastSlide',
  });
  assert.throws(() => replayEdits(conflict, history), /conflict/i);
});

test('changing one action trigger preserves an unrecognized action on the other trigger', async () => {
  for (const preserved of ['hlinkClick', 'hlinkHover']) {
    const source = await fixture();
    const parts = unzipSync(await pptx.savePresentation(source));
    const path = 'ppt/slides/slide1.xml';
    const xml = strFromU8(parts[path]);
    const action = `<a:${preserved} action="ppaction://macro?name=ExistingMacro" endSnd="1"><a:extLst><a:ext uri="test-extension"/></a:extLst></a:${preserved}>`;
    // Add to the first actual shape (not the slide's nonvisual group properties).
    parts[path] = strToU8(
      xml.replace(/(<p:nvSpPr>\s*<p:cNvPr[^>]*)(\/>)/, '$1>' + action + '</p:cNvPr>'),
    );
    assert.notEqual(strFromU8(parts[path]), xml);
    const p = await pptx.loadPresentation(zipSync(parts));
    const id = editorModel(p).slides[0].shapes[0].id;
    const next = { action: { kind: 'nextSlide' }, tooltip: null };
    recordEdit(p, {
      type: 'object-actions',
      slide: 0,
      ids: [id],
      link: preserved === 'hlinkClick' ? null : next,
      hoverLink: preserved === 'hlinkHover' ? null : next,
    });
    const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
    const after = pptx.getSlideXmlString(pptx.getSlides(saved)[0]);
    assert.ok(after.includes('ppaction://macro?name=ExistingMacro'));
    assert.ok(after.includes('endSnd="1"'));
    assert.ok(after.includes('test-extension'));
    const shape = editorModel(saved).slides[0].shapes[0];
    assert.deepEqual(preserved === 'hlinkClick' ? shape.hoverLink : shape.link, next);
  }
});

test('action sounds save, replay and undo together with links and validate before mutation', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const id = editorModel(p).slides[0].shapes[0].id;
  const bytes = Buffer.alloc(46);
  bytes.write('RIFF', 0);
  bytes.writeUInt32LE(38, 4);
  bytes.write('WAVEfmt ', 8);
  bytes.writeUInt32LE(16, 16);
  bytes.writeUInt16LE(1, 20);
  bytes.writeUInt16LE(1, 22);
  bytes.writeUInt32LE(8000, 24);
  bytes.writeUInt32LE(16000, 28);
  bytes.writeUInt16LE(2, 32);
  bytes.writeUInt16LE(16, 34);
  bytes.write('data', 36);
  bytes.writeUInt32LE(2, 40);
  const sounds = {
    click: { sound: { name: 'Chime.wav', base64: bytes.toString('base64') }, stopPrevious: false },
    hover: { sound: null, stopPrevious: true },
  };
  const command = {
    type: 'object-actions',
    slide: 0,
    ids: [id],
    link: { action: { kind: 'nextSlide' }, tooltip: null },
    hoverLink: null,
    actionSounds: sounds,
  };
  const history = emptyHistory();
  history.entries.push(recordEdit(p, command));
  history.cursor = 1;
  for (const target of [p, await pptx.loadPresentation(await pptx.savePresentation(p))])
    assert.deepEqual(editorModel(target).slides[0].shapes[0].actionSounds, sounds);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed).slides[0].shapes[0].actionSounds, sounds);
  recordEdit(p, { ...command, link: null });
  assert.equal(editorModel(p).slides[0].shapes[0].link, null);
  assert.deepEqual(editorModel(p).slides[0].shapes[0].actionSounds, sounds);
  const before = await pptx.savePresentation(p);
  assert.throws(
    () =>
      recordEdit(p, {
        ...command,
        actionSounds: {
          ...sounds,
          hover: { sound: { name: 'Broken', base64: 'AAAA' }, stopPrevious: false },
        },
      }),
    /WAV/,
  );
  assert.deepEqual(await pptx.savePresentation(p), before);
  history.cursor = 0;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, history);
  assert.equal(editorModel(undone).slides[0].shapes[0].actionSounds.click.sound, null);
  assert.equal(editorModel(undone).slides[0].shapes[0].actionSounds.hover.stopPrevious, false);
});

test('custom shows save, replay, undo and reject missing-slide or changed-source edits', async () => {
  const p = await fixture();
  pptx.addBlankSlide(p);
  const base = await pptx.savePresentation(p);
  const slides = editorModel(p).slides.map((slide) => slide.key);
  const customShows = [
    { id: 0, name: 'Selected slides', slides: [slides[1], slides[0], slides[1]] },
  ];
  const command = { type: 'custom-shows', slide: 0, customShows };
  const entry = recordEdit(p, command);
  assert.deepEqual(editorModel(p).customShows, customShows);
  assert.deepEqual(
    editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))).customShows,
    customShows,
  );
  const history = { ...emptyHistory(), entries: [entry], cursor: 1 };
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed).customShows, customShows);
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undone).customShows, []);
  const before = await pptx.savePresentation(p);
  assert.throws(() =>
    recordEdit(p, {
      ...command,
      customShows: [
        ...customShows,
        { id: 1, name: 'Missing', slides: ['/ppt/slides/missing.xml'] },
      ],
    }),
  );
  assert.deepEqual(await pptx.savePresentation(p), before);
  const changed = await pptx.loadPresentation(base);
  pptx.setCustomShows(changed, [
    { id: 7, name: 'Source edit', slides: [pptx.getSlides(changed)[0]] },
  ]);
  assert.throws(() => replayEdits(changed, history), /conflict|changed/i);
});

test('custom-show action edits replay, undo, and reject a missing show without changing either trigger', async () => {
  const p = await fixture();
  const slide = pptx.getSlides(p)[0];
  pptx.setCustomShows(p, [{ id: 9, name: 'Detail', slides: [slide] }]);
  const base = await pptx.savePresentation(p);
  const id = editorModel(p).slides[0].shapes[0].id;
  const link = { action: { kind: 'customShow', id: 9, showAndReturn: true }, tooltip: 'Detail' };
  const hoverLink = { action: { kind: 'customShow', id: 9, showAndReturn: false }, tooltip: null };
  const entry = recordEdit(p, { type: 'object-actions', slide: 0, ids: [id], link, hoverLink });
  const reloaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(reloaded).slides[0].shapes[0].link, link);
  assert.deepEqual(editorModel(reloaded).slides[0].shapes[0].hoverLink, hoverLink);
  const history = { ...emptyHistory(), entries: [entry], cursor: 1 };
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed).slides[0].shapes[0].link, link);
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...history, cursor: 0 });
  assert.equal(editorModel(undone).slides[0].shapes[0].link, null);
  const before = await pptx.savePresentation(p);
  assert.throws(
    () =>
      recordEdit(p, {
        type: 'object-actions',
        slide: 0,
        ids: [id],
        link: null,
        hoverLink: { ...link, action: { ...link.action, id: 999 } },
      }),
    /Custom show/,
  );
  assert.deepEqual(await pptx.savePresentation(p), before);
});

test('slideshow settings persist, replay, undo and detect changed settings', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const initial = editorModel(p).showProperties;
  const showProperties = { ...initial, slides: { kind: 'range', start: 1, end: 1 }, loop: true };
  const entry = recordEdit(p, { type: 'show-properties', slide: 0, showProperties });
  assert.deepEqual(editorModel(p).showProperties, showProperties);
  assert.deepEqual(
    editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))).showProperties,
    showProperties,
  );
  const history = { ...emptyHistory(), entries: [entry], cursor: 1 };
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed).showProperties, showProperties);
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undone).showProperties, initial);
  const changed = await pptx.loadPresentation(base);
  pptx.setSlideShowProperties(changed, { ...initial, useTimings: false });
  assert.throws(() => replayEdits(changed, history), /conflict|changed/i);
  const before = await pptx.savePresentation(p);
  assert.throws(() =>
    recordEdit(p, {
      type: 'show-properties',
      slide: 0,
      showProperties: { ...initial, slides: { kind: 'customShow', id: 999 } },
    }),
  );
  assert.deepEqual(await pptx.savePresentation(p), before);
});

test('slide advance timing persists, replays, undoes and detects conflicts', async () => {
  const p = await fixture();
  const base = await pptx.savePresentation(p);
  const advanceTiming = { advanceOnClick: false, advanceAfterMs: 2500 };
  const entry = recordEdit(p, { type: 'slide-advance', slide: 0, advanceTiming });
  const expected = { effect: 'none', ...advanceTiming };
  assert.deepEqual(editorModel(p).slides[0].transition, expected);
  assert.deepEqual(
    editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))).slides[0].transition,
    expected,
  );
  const history = { ...emptyHistory(), entries: [entry], cursor: 1 };
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed).slides[0].transition, expected);
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...history, cursor: 0 });
  assert.equal(editorModel(undone).slides[0].transition, null);
  const changed = await pptx.loadPresentation(base);
  pptx.setSlideTransition(pptx.getSlides(changed)[0], { effect: 'fade' });
  assert.throws(() => replayEdits(changed, history), /conflict|changed/i);
});

test('applying transitions to all slides replays, undoes and detects destination conflicts', async () => {
  const p = await fixture();
  const second = pptx.addBlankSlide(p);
  pptx.setSlideTransition(pptx.getSlides(p)[0], { effect: 'fade', advanceAfterMs: 1250 });
  pptx.setSlideTransition(second, { effect: 'push', direction: 'l' });
  const base = await pptx.savePresentation(p);
  const before = editorModel(p).slides.map((s) => s.transition);
  const entry = recordEdit(p, { type: 'transition-apply-all', slide: 0 });
  assert.deepEqual(
    editorModel(p).slides.map((s) => s.transition),
    [before[0], before[0]],
  );
  const history = { ...emptyHistory(), entries: [entry], cursor: 1 };
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(
    editorModel(replayed).slides.map((s) => s.transition),
    [before[0], before[0]],
  );
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...history, cursor: 0 });
  assert.deepEqual(
    editorModel(undone).slides.map((s) => s.transition),
    before,
  );
  const changed = await pptx.loadPresentation(base);
  pptx.setSlideTransition(pptx.getSlides(changed)[1], { effect: 'wipe' });
  assert.throws(() => replayEdits(changed, history), /conflict|changed/i);
});

test('transition effect edits persist, replay and undo without changing advance timing', async () => {
  const p = await fixture();
  pptx.setSlideAdvanceTiming(pptx.getSlides(p)[0], { advanceOnClick: false, advanceAfterMs: 1200 });
  const base = await pptx.savePresentation(p);
  const before = editorModel(p).slides[0].transition;
  const entry = recordEdit(p, {
    type: 'transition-effect',
    slide: 0,
    transitionEffect: { effect: 'wipe', direction: 'r' },
  });
  const expected = { ...before, effect: 'wipe', direction: 'r' };
  assert.deepEqual(editorModel(p).slides[0].transition, expected);
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(loaded).slides[0].transition, expected);
  const history = { ...emptyHistory(), entries: [entry], cursor: 1 };
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed).slides[0].transition, expected);
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undone).slides[0].transition, before);
});

test('animation duration preserves separate effects through save, undo and redo', async () => {
  const p = await fixture();
  const slide = pptx.getSlides(p)[0];
  pptx.setShapeAnimation(pptx.getSlideShapes(slide)[0], { effect: 'fadeIn', durationMs: 700 });
  pptx.setShapeAnimation(pptx.getSlideShapes(slide)[0], { effect: 'fadeOut', durationMs: 1200 });
  const base = await pptx.savePresentation(p);
  const before = editorModel(p).slides[0].animations;
  const id = before[1][0].timingId;
  const entry = recordEdit(p, {
    type: 'animation-duration',
    slide: 0,
    animationTiming: { id, durationMs: 2400 },
  });
  const expected = structuredClone(before);
  expected[1][0].durationMs = 2400;
  assert.deepEqual(editorModel(p).slides[0].animations, expected);
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(saved).slides[0].animations, expected);
  const history = { ...emptyHistory(), entries: [entry], cursor: 1 };
  const redone = await pptx.loadPresentation(base);
  replayEdits(redone, history);
  assert.deepEqual(editorModel(redone).slides[0].animations, expected);
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undone).slides[0].animations, before);
  const conflicting = await pptx.loadPresentation(base);
  pptx.setSlideAnimationDuration(pptx.getSlides(conflicting)[0], id, 1800);
  assert.throws(() => replayEdits(conflicting, history), /conflict|changed/i);
});

test('transition duration persists and participates in history', async () => {
  const p = await fixture();
  pptx.setSlideTransition(pptx.getSlides(p)[0], {
    effect: 'fade',
    durationMs: 500,
    advanceAfterMs: 3000,
  });
  const base = await pptx.savePresentation(p);
  const entry = recordEdit(p, {
    type: 'transition-duration',
    slide: 0,
    transitionDurationMs: 1750,
  });
  assert.equal(editorModel(p).slides[0].transition.durationMs, 1750);
  assert.equal(
    editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))).slides[0].transition
      .durationMs,
    1750,
  );
  const history = { ...emptyHistory(), entries: [entry], cursor: 1 };
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.equal(editorModel(replayed).slides[0].transition.durationMs, 1750);
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...history, cursor: 0 });
  assert.equal(editorModel(undone).slides[0].transition.durationMs, 500);
});

test('transition sound persists, replays, undoes and detects conflicts', async () => {
  const p = await fixture();
  const slide = pptx.getSlides(p)[0];
  pptx.setSlideTransition(slide, { effect: 'fade', durationMs: 800 });
  const base = await pptx.savePresentation(p);
  const bytes = Buffer.alloc(46);
  bytes.write('RIFF', 0);
  bytes.writeUInt32LE(38, 4);
  bytes.write('WAVE', 8);
  bytes.write('fmt ', 12);
  bytes.writeUInt32LE(16, 16);
  bytes.writeUInt16LE(1, 20);
  bytes.writeUInt16LE(1, 22);
  bytes.writeUInt32LE(8000, 24);
  bytes.writeUInt32LE(16000, 28);
  bytes.writeUInt16LE(2, 32);
  bytes.writeUInt16LE(16, 34);
  bytes.write('data', 36);
  bytes.writeUInt32LE(2, 40);
  const sound = { kind: 'play', name: 'Chime.wav', base64: bytes.toString('base64'), loop: true };
  const entry = recordEdit(p, { type: 'transition-sound', slide: 0, transitionSound: sound });
  assert.deepEqual(editorModel(p).slides[0].transitionSound, sound);
  assert.deepEqual(editorModel(p).slides[0].transition, { effect: 'fade', durationMs: 800 });
  const loaded = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(loaded).slides[0].transitionSound, sound);
  const history = { ...emptyHistory(), entries: [entry], cursor: 1 };
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed).slides[0].transitionSound, sound);
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undone).slides[0].transitionSound, { kind: 'none' });
  const changed = await pptx.loadPresentation(base);
  pptx.setSlideTransitionSound(pptx.getSlides(changed)[0], { kind: 'stop' });
  assert.throws(() => replayEdits(changed, history), /conflict|changed/i);
  const soundBase = await pptx.savePresentation(p);
  for (const command of [
    { type: 'transition-sound', slide: 0, transitionSound: { ...sound, loop: false } },
    { type: 'transition-apply-all', slide: 0 },
  ]) {
    const original = await pptx.loadPresentation(soundBase);
    const edit = recordEdit(original, command);
    const journal = { ...emptyHistory(), entries: [edit], cursor: 1 };
    const unchanged = await pptx.loadPresentation(soundBase);
    assert.doesNotThrow(() => replayEdits(unchanged, journal));
    const legacy = { ...edit };
    delete legacy.transitionSoundBefore;
    const legacyPresentation = await pptx.loadPresentation(soundBase);
    assert.doesNotThrow(() => replayEdits(legacyPresentation, { ...journal, entries: [legacy] }));
    const parts = unzipSync(soundBase);
    const media = Object.keys(parts).find((name) => name.endsWith('.wav'));
    assert.ok(media);
    parts[media][44] = 127;
    const externallyChanged = await pptx.loadPresentation(zipSync(parts));
    const before = editorModel(externallyChanged).slides[0].transitionSound;
    assert.throws(() => replayEdits(externallyChanged, journal), /conflict|changed/i);
    assert.deepEqual(editorModel(externallyChanged).slides[0].transitionSound, before);
  }
  for (const invalid of [
    { ...sound, base64: 'AAAA' },
    { ...sound, loop: 'true' },
    { kind: 'other' },
  ]) {
    const before = await pptx.savePresentation(p);
    assert.throws(() =>
      recordEdit(p, { type: 'transition-sound', slide: 0, transitionSound: invalid }),
    );
    assert.deepEqual(await pptx.savePresentation(p), before);
  }
  for (const kind of ['stop', 'none']) {
    recordEdit(p, { type: 'transition-sound', slide: 0, transitionSound: { kind } });
    assert.deepEqual(editorModel(p).slides[0].transitionSound, { kind });
  }
});

test('multi-object stacking preserves relative order at boundaries and through history', async () => {
  for (const nested of [false, true]) {
    for (const [order, selectedNames, expected] of [
      ['front', ['C', 'A'], ['B', 'D', 'A', 'C']],
      ['back', ['D', 'B'], ['B', 'D', 'A', 'C']],
      ['forward', ['C', 'B'], ['A', 'D', 'B', 'C']],
      ['forward', ['D', 'C'], ['A', 'B', 'C', 'D']],
      ['backward', ['C', 'B'], ['B', 'C', 'A', 'D']],
      ['backward', ['B', 'A'], ['A', 'B', 'C', 'D']],
    ]) {
      const p = pptx.createPresentation(),
        slide = pptx.addBlankSlide(p);
      const shapes = ['A', 'B', 'C', 'D'].map((name) =>
        pptx.addSlideShape(slide, {
          name,
          preset: 'rect',
          x: pptx.inches(1),
          y: pptx.inches(1),
          w: pptx.inches(2),
          h: pptx.inches(1),
        }),
      );
      if (nested) pptx.groupShapes(shapes);
      const base = await pptx.savePresentation(p),
        before = editorModel(p);
      const ids = selectedNames.map((name) =>
        pptx.getShapeId(shapes.find((s) => pptx.getShapeName(s) === name)),
      );
      const history = emptyHistory();
      history.entries.push(
        recordEdit(p, { type: 'order', slide: 0, ids, order, stableOrder: true }),
      );
      history.cursor = 1;
      const names = (model) =>
        (nested ? model.slides[0].shapes[0].children : model.slides[0].shapes).map((s) => s.name);
      assert.deepEqual(names(editorModel(p)), expected, `${nested}/${order}/${selectedNames}`);
      const replay = await pptx.loadPresentation(base);
      replayEdits(replay, history);
      assert.deepEqual(editorModel(replay), editorModel(p));
      history.cursor = 0;
      const undo = await pptx.loadPresentation(base);
      replayEdits(undo, history);
      assert.deepEqual(editorModel(undo), before);
    }
  }
});

test('duplicate preserves parent, stacking and slide-space offset inside transformed groups', async () => {
  const p = pptx.createPresentation(),
    slide = pptx.addBlankSlide(p);
  const children = ['A', 'B'].map((name, i) =>
    pptx.addSlideShape(slide, {
      name,
      preset: 'rect',
      x: pptx.inches(1 + i * 3),
      y: pptx.inches(2),
      w: pptx.inches(2),
      h: pptx.inches(1),
    }),
  );
  const group = pptx.groupShapes(children);
  pptx.setShapeRotation(group, 30);
  pptx.setShapeSize(group, pptx.inches(10), pptx.inches(3));
  const before = editorModel(p),
    base = await pptx.savePresentation(p);
  const history = emptyHistory();
  history.entries.push(
    recordEdit(p, {
      type: 'duplicate',
      slide: 0,
      ids: children.toReversed().map(pptx.getShapeId),
      duplicateInParent: true,
    }),
  );
  history.cursor = 1;
  const after = editorModel(p),
    parent = after.slides[0].shapes[0];
  assert.equal(after.slides[0].shapes.length, 1);
  assert.deepEqual(parent.bounds, before.slides[0].shapes[0].bounds);
  assert.deepEqual(
    parent.children.map((s) => s.name),
    ['A', 'B', 'A', 'B'],
  );
  for (let i = 0; i < 2; i++) {
    const original = parent.children[i],
      copy = parent.children[i + 2];
    assert.deepEqual(original, before.slides[0].shapes[0].children[i]);
    assert.notEqual(copy.id, original.id);
    const [a, b, c, d] = original.parentTransform;
    const dx = copy.bounds.x - original.bounds.x,
      dy = copy.bounds.y - original.bounds.y;
    assert.ok(Math.abs(a * dx + c * dy - pptx.pt(12)) < 3);
    assert.ok(Math.abs(b * dx + d * dy - pptx.pt(12)) < 3);
  }
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, history);
  assert.deepEqual(editorModel(replay), after);
  assert.deepEqual(editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))), after);
  history.cursor = 0;
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, history);
  assert.deepEqual(editorModel(undo), before);
});

test('snapshot paste releases nested ancestors into slide coordinates without editing source', async () => {
  for (const pasteOffset of [false, true]) {
    const p = pptx.createPresentation(),
      slide = pptx.addBlankSlide(p);
    const make = (name, x) =>
      pptx.addSlideShape(slide, {
        name,
        preset: 'rect',
        x: pptx.inches(x),
        y: pptx.inches(2),
        w: pptx.inches(2),
        h: pptx.inches(1),
      });
    const child = make('Selected child', 1),
      sibling = make('Sibling', 4);
    pptx.setShapeRotation(child, 23);
    const inner = pptx.groupShapes([child, sibling]);
    pptx.setShapeRotation(inner, 30);
    const outer = pptx.groupShapes([inner, make('Outer sibling', 8)]);
    pptx.setShapeRotation(outer, 15);
    const bounds = pptx.getShapeBoundsResolved(p, outer);
    pptx.setShapeSize(outer, pptx.emu(bounds.w * 2), pptx.emu(bounds.h * 2));
    const base = await pptx.savePresentation(p),
      before = editorModel(p);
    const original = before.slides[0].shapes[0].children[0].children[0];
    const history = emptyHistory();
    history.entries.push(
      recordEdit(p, {
        type: 'paste',
        slide: 0,
        source: 0,
        ids: [pptx.getShapeId(child)],
        snapshot: Buffer.from(base).toString('base64'),
        pasteSlideCoordinates: true,
        pasteOffset,
      }),
    );
    history.cursor = 1;
    const after = editorModel(p),
      copy = after.slides[0].shapes[1];
    assert.equal(after.slides[0].shapes.length, 2);
    assert.deepEqual(after.slides[0].shapes[0], before.slides[0].shapes[0]);
    assert.equal(copy.name, original.name);
    assert.equal(copy.rotation, 68);
    assert.equal(copy.parentTransform, undefined);
    const [a, b, c, d, e, f] = original.parentTransform;
    const x = original.bounds.x + original.bounds.w / 2;
    const y = original.bounds.y + original.bounds.h / 2;
    const offset = pasteOffset ? pptx.pt(12) : 0;
    assert.ok(Math.abs(copy.bounds.x + copy.bounds.w / 2 - (a * x + c * y + e + offset)) < 3);
    assert.ok(Math.abs(copy.bounds.y + copy.bounds.h / 2 - (b * x + d * y + f + offset)) < 3);
    assert.equal(copy.bounds.w, original.bounds.w * 2);
    assert.equal(copy.bounds.h, original.bounds.h * 2);
    const replay = await pptx.loadPresentation(base);
    replayEdits(replay, history);
    assert.deepEqual(editorModel(replay), after);
    assert.deepEqual(
      editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))),
      after,
    );
    history.cursor = 0;
    const undo = await pptx.loadPresentation(base);
    replayEdits(undo, history);
    assert.deepEqual(editorModel(undo), before);
  }
});

test('hidden group contents stay saved but leave the rendered slide and undo restores them', async () => {
  const p = await fixture();
  const slide = pptx.getSlides(p)[0];
  pptx.addSlideTextBox(slide, { x: 0, y: 0, w: 914400, h: 914400, text: 'Sibling' });
  const group = pptx.groupShapes(pptx.getSlideShapes(slide));
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const history = emptyHistory();
  history.entries.push(
    recordEdit(p, {
      type: 'update',
      slide: 0,
      ids: [pptx.getShapeId(group)],
      changes: { hidden: true },
    }),
  );
  history.cursor = 1;
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.equal(editorModel(saved).slides[0].shapes[0].hidden, true);
  assert.equal(editorModel(saved).slides[0].shapes[0].children[0].text, 'Original');
  assert.ok(!renderSlideToSvg(saved, pptx.getSlides(saved)[0]).includes('Original'));
  const redo = await pptx.loadPresentation(base);
  replayEdits(redo, history);
  assert.deepEqual(editorModel(redo), editorModel(saved));
  history.cursor = 0;
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, history);
  assert.deepEqual(editorModel(undo), before);
  assert.ok(renderSlideToSvg(undo, pptx.getSlides(undo)[0]).includes('Original'));
});

test('relative layer placement preserves selected order, parent and replay history', async () => {
  const p = await fixture();
  const slide = pptx.getSlides(p)[0];
  for (const text of ['Second', 'Third', 'Fourth'])
    pptx.addSlideTextBox(slide, { x: 0, y: 0, w: 914400, h: 914400, text });
  const children = pptx.getSlideShapes(slide);
  const ids = children.map(pptx.getShapeId);
  const group = pptx.groupShapes(children);
  const base = await pptx.savePresentation(p);
  const before = editorModel(p);
  const history = emptyHistory();
  history.entries.push(
    recordEdit(p, {
      type: 'order',
      slide: 0,
      ids: [ids[1], ids[0]],
      orderPlacement: { target: ids[3], side: 'after' },
    }),
  );
  history.cursor = 1;
  assert.deepEqual(pptx.getGroupChildren(group).map(pptx.getShapeId), [
    ids[2],
    ids[3],
    ids[0],
    ids[1],
  ]);
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  const redo = await pptx.loadPresentation(base);
  replayEdits(redo, history);
  assert.deepEqual(editorModel(redo), editorModel(saved));
  history.cursor = 0;
  const undo = await pptx.loadPresentation(base);
  replayEdits(undo, history);
  assert.deepEqual(editorModel(undo), before);
  assert.throws(() =>
    recordEdit(undo, {
      type: 'order',
      slide: 0,
      ids: [ids[0]],
      orderPlacement: { target: pptx.getShapeId(group), side: 'after' },
    }),
  );
  assert.deepEqual(editorModel(undo), before);
});

test('relative flips preserve independent flags and survive export, replay and undo', async () => {
  const p = await fixture();
  const id = editorModel(p).slides[0].shapes[0].id;
  recordEdit(p, { type: 'duplicate', slide: 0, ids: [id] });
  const second = pptx.getSlideShapes(pptx.getSlides(p)[0])[1];
  pptx.setShapeFlip(second, { horizontal: true, vertical: true });
  const base = await pptx.savePresentation(p);
  const original = editorModel(p).slides[0].shapes;
  const ids = original.map((shape) => shape.id);
  const history = emptyHistory();
  for (const direction of ['horizontal', 'vertical']) {
    history.entries.push(
      recordEdit(p, { type: 'update', slide: 0, ids, changes: { flipToggle: direction } }),
    );
    history.cursor++;
  }
  const result = editorModel(p);
  assert.deepEqual(
    result.slides[0].shapes.map((shape) => shape.flip),
    [
      { horizontal: true, vertical: true },
      { horizontal: false, vertical: false },
    ],
  );
  assert.deepEqual(
    result.slides[0].shapes.map((shape) => shape.bounds),
    original.map((shape) => shape.bounds),
  );
  const saved = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.deepEqual(editorModel(saved), result);
  const replayed = await pptx.loadPresentation(base);
  replayEdits(replayed, history);
  assert.deepEqual(editorModel(replayed), result);
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, { ...history, cursor: 0 });
  assert.deepEqual(editorModel(undone).slides[0].shapes, original);
  assert.throws(
    () => recordEdit(p, { type: 'update', slide: 0, ids, changes: { flipToggle: 'diagonal' } }),
    /Invalid flip direction/,
  );
});

test('regroup restores former members after edits and history replay', async () => {
  const p = await fixture();
  const slide = pptx.getSlides(p)[0];
  pptx.addSlideTextBox(slide, {
    x: pptx.inches(6),
    y: pptx.inches(1),
    w: pptx.inches(2),
    h: pptx.inches(1),
    text: 'Second',
  });
  const base = await pptx.savePresentation(p);
  const history = emptyHistory();
  const edit = (command) => {
    history.entries.push(recordEdit(p, { slide: 0, ...command }));
    history.cursor++;
  };
  const ids = editorModel(p).slides[0].shapes.map((s) => s.id);
  edit({ type: 'group', ids });
  edit({ type: 'ungroup', ids: [editorModel(p).slides[0].shapes[0].id] });
  edit({ type: 'update', ids: [ids[0]], changes: { rotationDelta: 15 } });
  const before = editorModel(p);
  assert.deepEqual(before.slides[0].shapes[0].regroupIds, ids);
  const roundTrip = await pptx.loadPresentation(await pptx.savePresentation(p));
  assert.equal(editorModel(roundTrip).slides[0].shapes[0].regroupIds, undefined);
  restoreEditorSession(roundTrip, before);
  assert.deepEqual(editorModel(roundTrip), before);
  recordEdit(roundTrip, { type: 'regroup', slide: 0, ids: [ids[1]] });
  edit({ type: 'regroup', ids: [ids[1]] });
  const grouped = editorModel(p);
  assert.deepEqual(editorModel(roundTrip), grouped);
  assert.equal(grouped.slides[0].shapes.length, 1);
  assert.deepEqual(
    grouped.slides[0].shapes[0].children.map((s) => s.id),
    ids,
  );
  assert.equal(grouped.slides[0].shapes[0].children[0].rotation, 15);
  assert.deepEqual(
    editorModel(await pptx.loadPresentation(await pptx.savePresentation(p))),
    grouped,
  );
  const replay = await pptx.loadPresentation(base);
  replayEdits(replay, history);
  assert.deepEqual(editorModel(replay), grouped);
  history.cursor--;
  const undone = await pptx.loadPresentation(base);
  replayEdits(undone, history);
  assert.deepEqual(editorModel(undone), before);
});
