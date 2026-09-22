// Resizing and rotating a child *inside* a transformed group, by dragging the
// handles the way a person does.
//
// The group-editing tests move a child. Moving only needs the pointer delta
// mapped into the group's space; resizing has to keep an anchor still on the
// slide while the child's own box changes, and rotating has to turn the child
// about a centre that does not move. Neither is covered by a move.
//
// The parents here are deliberately awkward: an oblique angle rather than a
// quarter turn, a scale that differs between the axes, and — in the second
// scenario — a reflection on top of both. What is asserted is read twice over,
// from the screen and from the saved file, and never from the product's own
// transform code.

import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getGroupChildren,
  getGroupTransform,
  getShapeBounds,
  getShapeFlip,
  getShapeId,
  getShapeRotation,
  getSlideShapes,
  getSlides,
  groupShapes,
  inches,
  loadPresentation,
  savePresentation,
  setShapeBounds,
  setShapeFlip,
  setShapeRotation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

/** How far a screen reading may drift before it stops meaning what it says. */
const SCREEN_SLACK = 4;

const centreOf = (box) => ({ x: box.x + box.width / 2, y: box.y + box.height / 2 });
const apart = (a, b) => Math.hypot(a.x - b.x, a.y - b.y);
const minus = (a, b) => ({ x: a.x - b.x, y: a.y - b.y });
const scaled = (v, k) => ({ x: v.x * k, y: v.y * k });
const plus = (a, b) => ({ x: a.x + b.x, y: a.y + b.y });

/**
 * The map from the space the child's own geometry is written in — the group's
 * child coordinates — to the screen, worked out from where three of the child's
 * handles are drawn and the turn the child itself carries.
 *
 * Built here rather than asked of the editor: an expectation taken from the
 * code under test only says the code agrees with itself.
 *
 * The handles sit at the corners of the child's box *after* its own rotation,
 * so if `M` is the map being looked for and `R` the child's own turn, the
 * measured edge vectors are `M·R·(w,0)` and `M·R·(0,h)`. Undoing `R` gives
 * `M`'s own two columns.
 */
const groupSpaceBasis = ({ nw, ne, sw }, { w, h, rotation }) => {
  const radians = (rotation * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const alongW = scaled(minus(ne, nw), 1 / w);
  const alongH = scaled(minus(sw, nw), 1 / h);
  return {
    x: minus(scaled(alongW, cos), scaled(alongH, sin)),
    y: plus(scaled(alongW, sin), scaled(alongH, cos)),
  };
};

/** Where a screen point stands, in the child's space, relative to `origin`. */
const intoGroupSpace = (basis, origin, point) => {
  const determinant = basis.x.x * basis.y.y - basis.x.y * basis.y.x;
  const offset = minus(point, origin);
  return {
    x: (offset.x * basis.y.y - offset.y * basis.y.x) / determinant,
    y: (offset.y * basis.x.x - offset.x * basis.x.y) / determinant,
  };
};

const degrees = (v) => (Math.atan2(v.y, v.x) * 180) / Math.PI;
/** The smallest turn between two angles, so 359° and 1° are two apart. */
const turnBetween = (a, b) => ((((a - b) % 360) + 540) % 360) - 180;

/**
 * Builds a deck whose single group holds two children, then gives the group an
 * oblique turn, a scale that differs between the axes, and optionally a
 * reflection. `setShapeBounds` on a group moves and resizes the group's own
 * frame while its child coordinate space stays as it was, which is exactly the
 * scale a deck carries when someone has dragged a group's corner.
 */
const skew = (group) => {
  const frame = getShapeBounds(group);
  setShapeBounds(group, {
    x: frame.x,
    y: frame.y,
    w: Math.round(frame.w * 1.8),
    h: Math.round(frame.h * 0.55),
  });
};

/**
 * The words the editor is driven by. One scenario runs in Japanese: the handles
 * are found by their accessible names, so a translated name that never reached
 * a handle would go unnoticed if every test spoke English.
 */
const WORDS = {
  en: (key) => key,
  ja: (key) =>
    ({
      'Editing group': 'グループを編集中',
      'Saved to this project': 'このプロジェクトに保存済み',
      'Undo (Ctrl+Z)': '元に戻す (Ctrl+Z)',
      'Redo (Ctrl+Y)': 'やり直し (Ctrl+Y)',
      Rotate: '回転',
      'Resize nw': 'サイズ変更（左上）',
      'Resize ne': 'サイズ変更（右上）',
      'Resize sw': 'サイズ変更（左下）',
      'Resize se': 'サイズ変更（右下）',
    })[key],
};

const buildDeck = ({ reflected, childRotation, nested }) => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const first = addSlideTextBox(slide, {
    x: inches(2),
    y: inches(2),
    w: inches(2),
    h: inches(1),
    text: '子',
  });
  const second = addSlideTextBox(slide, {
    x: inches(5),
    y: inches(2.5),
    w: inches(1.5),
    h: inches(1),
    text: 'Sibling',
  });
  if (childRotation) setShapeRotation(first, childRotation);
  let group = groupShapes([first, second]);
  skew(group);
  setShapeRotation(group, 37);
  if (reflected) setShapeFlip(group, { horizontal: true });
  const frames = [group];
  let depth = 1;
  if (nested) {
    // A second frame around the first, turned the other way and scaled again,
    // so the child's own space is two transforms down rather than one.
    const outside = addSlideTextBox(slide, {
      x: inches(0.5),
      y: inches(5),
      w: inches(1),
      h: inches(1),
      text: 'Outer',
    });
    group = groupShapes([group, outside]);
    skew(group);
    setShapeRotation(group, -23);
    frames.push(group);
    depth = 2;
  }
  return {
    pres,
    depth,
    childId: getShapeId(first),
    siblingId: getShapeId(second),
    // Every frame between the child and the slide. All of them have to come
    // through the child's edits untouched, inner coordinate space included.
    frameIds: frames.map(getShapeId),
  };
};

const openEditor = async (dir, pres, language) => {
  const source = join(dir, 'source.pptx');
  const file = join(dir, 'deck.tsx');
  await writeFile(source, await savePresentation(pres));
  await writeFile(
    file,
    `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';` +
      `export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
  );
  const preview = await startPreview(file);
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
  const errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto(preview.url);
  const word = WORDS[language];
  const reopen = async () => {
    await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
    const frame = page.frameLocator('#editor-frame');
    // The chosen language outlives a reload, so this asks for it only when the
    // editor is not already speaking it.
    const picker = frame.locator('.lang select');
    await picker.waitFor();
    if ((await picker.inputValue()) !== language) await picker.selectOption(language);
    await frame.getByText(word('Saved to this project'), { exact: true }).waitFor();
    return frame;
  };
  const editor = await reopen();
  return { preview, browser, page, editor, errors, word, reopen };
};

/**
 * The shapes the deck on disk holds. `getSlideShapes` already descends into
 * groups, so a group and its children are all in the one list.
 */
const savedShapes = async (preview) => {
  const pres = await loadPresentation(
    new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
  );
  return getSlideShapes(getSlides(pres)[0]);
};

/**
 * What the saved file says about one shape. For a group that is more than its
 * own frame: the inner coordinate space its children are written against
 * (`chOff` / `chExt`) moves the children on the slide without any of their own
 * numbers changing, and so does reordering them — so both are part of what
 * "the group is untouched" has to mean.
 */
const geometryOf = async (preview, id) => {
  const shape = (await savedShapes(preview)).find((s) => getShapeId(s) === id);
  assert.ok(shape, `shape ${id} is still in the deck`);
  return {
    ...getShapeBounds(shape),
    rotation: getShapeRotation(shape),
    flip: getShapeFlip(shape),
    frame: getGroupTransform(shape),
    children: getGroupChildren(shape).map(getShapeId),
  };
};

/**
 * Waits for the deck on disk to satisfy `done`, rather than for the save
 * indicator — which is already showing "saved" from the edit before this one,
 * so waiting on it can step straight past the write being checked.
 */
const settles = async (preview, id, done, what) => {
  const deadline = Date.now() + 20000;
  let last;
  for (;;) {
    last = await geometryOf(preview, id);
    if (done(last)) return last;
    if (Date.now() > deadline) {
      assert.fail(`${what}: the saved geometry never got there, last was ${JSON.stringify(last)}`);
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
};

const dragBy = async (page, from, to) => {
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 12 });
  await page.mouse.up();
};

/** Steps down into the group — one step per frame — and selects the child. */
const selectChild = async (editor, depth, word) => {
  for (let step = 0; step < depth; step += 1) {
    await editor.locator('.hit').first().dblclick();
    await editor.getByText(word('Editing group'), { exact: true }).waitFor();
  }
  assert.equal(await editor.locator('.hit').count(), 2);
  await editor.locator('.hit').first().click();
};

const scenario = (name, options) =>
  test(
    `a child of a ${name} group resizes and rotates from its own handles`,
    { timeout: 120000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-group-handles-'));
      let session;
      try {
        const { pres, depth, childId, siblingId, frameIds } = buildDeck(options);
        session = await openEditor(dir, pres, options.language ?? 'en');
        const { preview, page, editor, errors, word, reopen } = session;

        const before = await geometryOf(preview, childId);
        const siblingBefore = await geometryOf(preview, siblingId);
        const framesBefore = [];
        for (const id of frameIds) framesBefore.push(await geometryOf(preview, id));
        assert.equal(before.rotation, options.childRotation ?? 0);
        // Nothing below is worth reading if the parents were never transformed
        // in the first place.
        for (const frame of framesBefore) {
          assert.ok(frame.frame, 'the group carries an explicit transform');
          assert.notEqual(frame.frame.outer.w, frame.frame.inner.w);
          assert.notEqual(frame.frame.outer.h, frame.frame.inner.h);
        }
        /** Every frame around the child is exactly as it was. */
        const framesUntouched = async (when) => {
          for (const [at, id] of frameIds.entries()) {
            assert.deepEqual(
              await geometryOf(preview, id),
              framesBefore[at],
              `frame ${id} ${when}`,
            );
          }
          assert.deepEqual(await geometryOf(preview, siblingId), siblingBefore, `sibling ${when}`);
        };

        await selectChild(editor, depth, word);
        const handle = (name) => editor.getByRole('button', { name: word(name), exact: true });
        const button = (name) => editor.getByTitle(word(name), { exact: true });
        await handle('Resize se').waitFor();

        // --- Resizing ------------------------------------------------------
        const anchorBefore = centreOf(await handle('Resize nw').boundingBox());
        const grip = centreOf(await handle('Resize se').boundingBox());
        const drop = { x: grip.x + 110, y: grip.y + 70 };
        await dragBy(page, grip, drop);

        const resized = await settles(
          preview,
          childId,
          (now) => now.w !== before.w || now.h !== before.h,
          'resize',
        );
        // The handle went where the pointer went. This is the whole of the
        // mapping — screen to the group's own space and back — checked without
        // asking the code that did it what it thinks the answer is.
        assert.ok(
          apart(centreOf(await handle('Resize se').boundingBox()), drop) < SCREEN_SLACK * 2,
          'the dragged corner follows the pointer',
        );
        // And the corner opposite it did not move: that is what "resize" means
        // as opposed to "move and resize".
        assert.ok(
          apart(centreOf(await handle('Resize nw').boundingBox()), anchorBefore) < SCREEN_SLACK,
          'the opposite corner stays put',
        );
        // The south-east handle moves that corner and leaves the north-west
        // one where it is, which in the child's own space is its origin. Which
        // way round the box grows is not asserted: under an oblique turn, an
        // uneven scale and a reflection, a drag down and to the right on screen
        // need not be one down and to the right in the space the child lives
        // in — and the screen readings above are what says it went the right
        // way there.
        if ((options.childRotation ?? 0) === 0) {
          // With no turn of its own, keeping the north-west corner still is
          // keeping the child's own origin still.
          assert.equal(resized.x, before.x, 'the child keeps its origin');
          assert.equal(resized.y, before.y, 'the child keeps its origin');
        } else {
          // With one, the box is drawn about its centre, so holding a corner
          // still on the slide has to move the origin. The screen readings
          // above are what says it moved to the right place.
          assert.ok(
            resized.x !== before.x || resized.y !== before.y,
            'a turned child re-anchors as it resizes',
          );
        }
        assert.ok(
          Math.abs(resized.w - before.w) > inches(0.2),
          `width barely moved: ${resized.w} vs ${before.w}`,
        );
        assert.ok(
          Math.abs(resized.h - before.h) > inches(0.2),
          `height barely moved: ${resized.h} vs ${before.h}`,
        );
        assert.equal(resized.rotation, before.rotation);
        // Neither the shape beside it nor any frame around it is touched.
        await framesUntouched('after the resize');

        await button('Undo (Ctrl+Z)').click();
        await settles(preview, childId, (now) => now.w === before.w && now.h === before.h, 'undo');
        await button('Redo (Ctrl+Y)').click();
        await settles(
          preview,
          childId,
          (now) => now.w === resized.w && now.h === resized.h,
          'redo',
        );

        // --- Rotating ------------------------------------------------------
        await editor.locator('.hit').first().click();
        await handle('Rotate').waitFor();
        const corners = {
          nw: centreOf(await handle('Resize nw').boundingBox()),
          ne: centreOf(await handle('Resize ne').boundingBox()),
          sw: centreOf(await handle('Resize sw').boundingBox()),
        };
        const centreBefore = centreOf(await editor.locator('.hit.selected').boundingBox());
        const basis = groupSpaceBasis(corners, resized);
        const knob = centreOf(await handle('Rotate').boundingBox());
        const radius = apart(knob, centreBefore);
        assert.ok(radius > 10, `the rotate knob sits off the shape: ${radius}`);
        // A quarter turn clockwise about the centre, on screen. In screen
        // coordinates y runs downwards, so turning (dx, dy) clockwise gives
        // (-dy, dx).
        const reach = minus(knob, centreBefore);
        const target = plus(centreBefore, { x: -reach.y, y: reach.x });
        await dragBy(page, knob, target);

        const turned = await settles(
          preview,
          childId,
          (now) => now.rotation !== resized.rotation,
          'rotate',
        );
        // How far round the pointer actually went, measured in the space the
        // stored angle is written in. A quarter turn on screen is not a quarter
        // turn there — the group scales the axes differently — so the swept
        // angle is worked out from the screen points rather than assumed.
        const swept = turnBetween(
          degrees(intoGroupSpace(basis, centreBefore, target)),
          degrees(intoGroupSpace(basis, centreBefore, knob)),
        );
        // A swept angle near zero would make the comparison below true of
        // almost any rotation, so it is worth knowing it is a real turn.
        assert.ok(Math.abs(swept) > 20, `the drag barely turned anything: ${swept}°`);
        const expected = resized.rotation + swept;
        assert.ok(
          Math.abs(turnBetween(turned.rotation, expected)) <= 3,
          `turned to ${turned.rotation}°, expected about ${expected.toFixed(1)}° ` +
            `(swept ${swept.toFixed(1)}° in the child's own space)`,
        );
        // Turning is about the centre, so the centre is where it was.
        assert.ok(
          apart(centreOf(await editor.locator('.hit.selected').boundingBox()), centreBefore) <
            SCREEN_SLACK * 3,
          'the centre of rotation stays put',
        );
        // The box itself is unchanged — only the angle it is drawn at.
        assert.equal(turned.w, resized.w);
        assert.equal(turned.h, resized.h);
        assert.equal(turned.x, resized.x);
        assert.equal(turned.y, resized.y);
        await framesUntouched('after the rotation');

        // --- Resizing something that is now turned twice over --------------
        // The child's own angle and everything the group does above it are both
        // in the way of the anchor now.
        //
        // The drag runs outward along the measured north-west → south-east
        // diagonal. Under an affine parent that is a straight enlargement of
        // both of the child's own dimensions whatever the parent does, which
        // keeps the box clear of the editor's minimum size — a clamp there
        // would stop the corner under the pointer for a good reason and say
        // nothing about the mapping this is checking.
        const anchorTurned = centreOf(await handle('Resize nw').boundingBox());
        const gripTurned = centreOf(await handle('Resize se').boundingBox());
        const dropTurned = plus(gripTurned, scaled(minus(gripTurned, anchorTurned), 0.4));
        await dragBy(page, gripTurned, dropTurned);
        const resizedAgain = await settles(
          preview,
          childId,
          (now) => now.w !== turned.w || now.h !== turned.h,
          'resize after rotating',
        );
        const gripLanded = centreOf(await handle('Resize se').boundingBox());
        assert.ok(
          apart(gripLanded, dropTurned) < SCREEN_SLACK * 2,
          `the dragged corner follows the pointer after a turn: landed ${JSON.stringify(gripLanded)}, ` +
            `dropped at ${JSON.stringify(dropTurned)}, from ${JSON.stringify(gripTurned)}; ` +
            `saved ${JSON.stringify(resizedAgain)} was ${JSON.stringify(turned)}`,
        );
        assert.ok(
          apart(centreOf(await handle('Resize nw').boundingBox()), anchorTurned) < SCREEN_SLACK,
          'the opposite corner stays put after a turn',
        );
        assert.equal(resizedAgain.rotation, turned.rotation);
        // Outward along the diagonal, so both dimensions grew — and by the same
        // proportion, since that drag is a pure enlargement about the anchor.
        assert.ok(resizedAgain.w > turned.w, `width shrank: ${resizedAgain.w} vs ${turned.w}`);
        assert.ok(resizedAgain.h > turned.h, `height shrank: ${resizedAgain.h} vs ${turned.h}`);
        assert.ok(
          Math.abs(resizedAgain.w / turned.w - resizedAgain.h / turned.h) < 0.05,
          `grew unevenly: ${resizedAgain.w / turned.w} vs ${resizedAgain.h / turned.h}`,
        );
        await framesUntouched('after the second resize');

        // --- Taking it back and putting it forward again -------------------
        await button('Undo (Ctrl+Z)').click();
        await settles(
          preview,
          childId,
          (now) => now.w === turned.w && now.h === turned.h,
          'undo the second resize',
        );
        await button('Undo (Ctrl+Z)').click();
        await settles(
          preview,
          childId,
          (now) => now.rotation === resized.rotation,
          'undo the rotation',
        );
        await button('Redo (Ctrl+Y)').click();
        const redone = await settles(
          preview,
          childId,
          (now) => now.rotation === turned.rotation,
          'redo the rotation',
        );

        // --- What a reload shows -------------------------------------------
        // Everything checked so far was the document in hand. This is the file.
        await page.reload();
        await reopen();
        assert.deepEqual(await geometryOf(preview, childId), redone);
        await framesUntouched('after a reload');

        assert.deepEqual(errors, []);
      } finally {
        await session?.browser.close();
        await session?.preview.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );

scenario('turned and unevenly scaled', { reflected: false, childRotation: 0 });
scenario('turned, unevenly scaled and reflected', {
  reflected: true,
  childRotation: 0,
  language: 'ja',
});
// The child carries a turn of its own, so the anchor has to survive both it and
// everything the group does above it.
scenario('turned and unevenly scaled, around a turned child', {
  reflected: false,
  childRotation: 20,
});
scenario('turned twice over, one frame inside another', { nested: true, childRotation: 20 });

// Dragged far enough, a corner arrives at the one opposite it and then goes
// past. The editor holds a floor there. Inside a transformed group that floor
// is worth its own test: the pointer and the box part company, which is the one
// place the checks above would be wrong to insist they agree.
test(
  'a child driven past its own opposite corner stops instead of turning inside out',
  {
    timeout: 120000,
  },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-group-handles-'));
    let session;
    try {
      const { pres, depth, childId, frameIds } = buildDeck({ reflected: true, childRotation: 20 });
      session = await openEditor(dir, pres, 'en');
      const { preview, page, editor, errors, word } = session;

      const before = await geometryOf(preview, childId);
      const framesBefore = [];
      for (const id of frameIds) framesBefore.push(await geometryOf(preview, id));
      await selectChild(editor, depth, word);
      const handle = (name) => editor.getByRole('button', { name, exact: true });
      await handle('Resize se').waitFor();

      // Inward along the diagonal and a long way past the anchor.
      const shoved = async () => {
        const anchor = centreOf(await handle('Resize nw').boundingBox());
        const grip = centreOf(await handle('Resize se').boundingBox());
        await dragBy(page, grip, plus(grip, scaled(minus(anchor, grip), 4)));
        return anchor;
      };
      const anchor = await shoved();
      const floored = await settles(
        preview,
        childId,
        (now) => now.w !== before.w || now.h !== before.h,
        'resize past the anchor',
      );
      assert.ok(floored.w > 0 && floored.h > 0, `the box inverted: ${JSON.stringify(floored)}`);
      assert.ok(floored.w < before.w / 2 && floored.h < before.h / 2, 'the drag did shrink it');
      assert.ok(
        apart(centreOf(await handle('Resize nw').boundingBox()), anchor) < SCREEN_SLACK,
        'the opposite corner stays put even at the floor',
      );

      // Shoving it again changes nothing, which is what says it stopped at a
      // floor rather than merely arriving somewhere small this once.
      await shoved();
      await new Promise((resolve) => setTimeout(resolve, 1000));
      assert.deepEqual(await geometryOf(preview, childId), floored);
      for (const [at, id] of frameIds.entries()) {
        assert.deepEqual(await geometryOf(preview, id), framesBefore[at]);
      }
      assert.deepEqual(errors, []);
    } finally {
      await session?.browser.close();
      await session?.preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
