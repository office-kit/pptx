import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'Group child canvas selection, drag and resize preserve parent coordinates',
  { timeout: 90000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-group-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={2} y={2} width={3} height={1}>First child</Text><Text x={7} y={4} width={2} height={1}>Second child</Text></Slide></Presentation>`,
    );
    const proc = spawn(process.execPath, [
      fileURLToPath(new URL('../../dist/cli.mjs', import.meta.url)),
      'dev',
      file,
      '--port',
      '0',
    ]);
    let browser;
    try {
      const url = await new Promise((resolve, reject) => {
        let output = '';
        proc.stdout.on('data', (data) => {
          output += data;
          const match = output.match(/Preview: (http:\/\/\S+)/);
          if (match) resolve(match[1]);
        });
        proc.stderr.on('data', (data) => process.stderr.write(data));
        proc.on('error', reject);
        proc.on('exit', (code) => reject(new Error('Server exited: ' + code)));
      });
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
      await page.goto(url);
      await page.locator('.shape-hit').first().waitFor();

      const state = () => page.evaluate(() => window.office.getState());
      const settled = async (revision) =>
        page.waitForFunction((rev) => {
          const s = window.office.getState();
          return (
            s.revision > rev &&
            !s.building &&
            !document.querySelector('[data-edit="undo"]').disabled
          );
        }, revision);
      await page.locator('.shape-hit').first().click();
      const initialIds = (await state()).editor.slides[0].shapes.map((s) => s.id);
      for (const [key, id] of [
        ['Shift+Tab', initialIds.at(-1)],
        ['Tab', initialIds[0]],
      ]) {
        await page.keyboard.press('Escape');
        await page.keyboard.press(key);
        assert.equal(
          await page.locator('.shape-hit.selected').getAttribute('data-shape-id'),
          String(id),
        );
      }
      await page.keyboard.press('Meta+a');
      let revision = (await state()).revision;
      await page.keyboard.press('Meta+Alt+g');
      await settled(revision);
      revision = (await state()).revision;
      await page.keyboard.press('Alt+ArrowRight');
      await settled(revision);
      const original = (await state()).editor.slides[0].shapes[0];
      assert.equal(original.rotation, 15);
      const child = original.children[0];
      const sibling = original.children[1];
      const canvas = await page.locator('#slide').boundingBox();
      const scale = canvas.width / (await state()).editor.width;
      const [a, b, c, d, e, f] = child.parentTransform;
      const centerX = child.bounds.x + child.bounds.w / 2;
      const centerY = child.bounds.y + child.bounds.h / 2;
      const x = canvas.x + (a * centerX + c * centerY + e) * scale;
      const y = canvas.y + (b * centerX + d * centerY + f) * scale;
      revision = (await state()).revision;
      // Clicking within an already selected group must enter that child without an edit.
      await page.mouse.click(x, y);
      assert.equal(
        await page.locator('.shape-hit.selected').getAttribute('data-shape-id'),
        String(child.id),
      );
      assert.equal((await state()).revision, revision);
      for (const [key, id] of [
        ['Tab', sibling.id],
        ['Tab', child.id],
        ['Shift+Tab', sibling.id],
        ['Shift+Tab', child.id],
      ]) {
        await page.keyboard.press(key);
        assert.equal(
          await page.locator('.shape-hit.selected').getAttribute('data-shape-id'),
          String(id),
        );
      }
      assert.equal((await state()).revision, revision);
      // Start at another point to avoid the text double-click gesture.
      await page.waitForTimeout(450);
      revision = (await state()).revision;
      await page.mouse.move(x, y);
      await page.mouse.down();
      await page.mouse.move(x + 40, y + 25, { steps: 6 });
      await page.mouse.up();
      await settled(revision);
      const movedGroup = (await state()).editor.slides[0].shapes[0];
      const moved = movedGroup.children[0];
      const dx = moved.bounds.x - child.bounds.x,
        dy = moved.bounds.y - child.bounds.y;
      assert.ok(Math.abs((a * dx + c * dy) * scale - 40) < 1);
      assert.ok(Math.abs((b * dx + d * dy) * scale - 25) < 1);
      assert.deepEqual(movedGroup.bounds, original.bounds);
      assert.deepEqual(movedGroup.children[1], sibling);
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], original);
      // A descendant snaps to its sibling in slide space while its parent stays fixed.
      const siblingX =
        a * (sibling.bounds.x + sibling.bounds.w / 2) +
        c * (sibling.bounds.y + sibling.bounds.h / 2) +
        e;
      const targetX = canvas.x + siblingX * scale;
      await page.waitForTimeout(450);
      revision = (await state()).revision;
      await page.mouse.move(x, y);
      await page.mouse.down();
      await page.mouse.move(targetX - 3, y + 17, { steps: 6 });
      assert.ok((await page.locator('.smart-guide[data-axis="x"]').count()) > 0);
      await page.mouse.up();
      await settled(revision);
      const snappedParent = (await state()).editor.slides[0].shapes[0];
      const snappedChild = snappedParent.children[0];
      const snappedX =
        a * (snappedChild.bounds.x + snappedChild.bounds.w / 2) +
        c * (snappedChild.bounds.y + snappedChild.bounds.h / 2) +
        e;
      assert.ok(Math.abs(snappedX - siblingX) * scale < 1);
      assert.deepEqual(snappedParent.bounds, original.bounds);
      assert.deepEqual(snappedParent.children[1], sibling);
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], original);
      const handle = await page.locator('.shape-hit.selected [data-handle="se"]').boundingBox();
      revision = (await state()).revision;
      await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2);
      await page.mouse.down();
      // Move along both parent-local axes after projection to the rotated slide.
      const localDx = 250000,
        localDy = 180000;
      await page.mouse.move(
        handle.x + handle.width / 2 + (a * localDx + c * localDy) * scale,
        handle.y + handle.height / 2 + (b * localDx + d * localDy) * scale,
        { steps: 6 },
      );
      await page.mouse.up();
      await settled(revision);
      const resizedGroup = (await state()).editor.slides[0].shapes[0];
      const resized = resizedGroup.children[0];
      assert.ok(Math.abs(resized.bounds.w - child.bounds.w - localDx) * scale < 1);
      assert.ok(Math.abs(resized.bounds.h - child.bounds.h - localDy) * scale < 1);
      assert.equal(resized.bounds.x, child.bounds.x);
      assert.equal(resized.bounds.y, child.bounds.y);
      assert.deepEqual(resizedGroup.children[1], sibling);
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], original);
      await page.locator('[data-edit="copy"]').click();
      await page.waitForFunction(() => !document.querySelector('[data-edit="paste"]').disabled);
      revision = (await state()).revision;
      await page.locator('[data-edit="paste"]').click();
      await settled(revision);
      const pastedShapes = (await state()).editor.slides[0].shapes;
      assert.equal(pastedShapes.length, 2);
      assert.deepEqual(pastedShapes[0], original);
      const pasted = pastedShapes[1];
      assert.equal(pasted.text, child.text);
      assert.equal(pasted.rotation, original.rotation);
      assert.equal(pasted.parentTransform, undefined);
      assert.ok(
        Math.abs(pasted.bounds.x + pasted.bounds.w / 2 - (a * centerX + c * centerY + e + 152400)) <
          3,
      );
      assert.ok(
        Math.abs(pasted.bounds.y + pasted.bounds.h / 2 - (b * centerX + d * centerY + f + 152400)) <
          3,
      );
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes, [original]);
      await page.reload();
      await page.locator('.shape-hit').first().waitFor();
      assert.deepEqual((await state()).editor.slides[0].shapes[0], original);
      await page.locator('[data-edit="selection"]:visible').first().click();
      await page.locator('#selection-pane [aria-expanded="false"]').click();
      await page.locator(`#selection-pane [data-selection-id="${sibling.id}"]`).click();
      const selectionRevision = (await state()).revision;
      const childRow = page.locator(`#selection-pane [data-selection-id="${child.id}"]`);
      await childRow.click({ modifiers: ['Meta'] });
      assert.equal(await page.locator('.shape-hit.selected').count(), 2);
      assert.equal(await childRow.getAttribute('aria-pressed'), 'true');
      await childRow.click({ modifiers: ['Meta'] });
      assert.equal(await page.locator('.shape-hit.selected').count(), 1);
      assert.equal(await childRow.getAttribute('aria-pressed'), 'false');
      await page
        .locator(`#selection-pane [data-selection-id="${original.id}"]`)
        .click({ modifiers: ['Meta'] });
      assert.equal(await page.locator('.shape-hit.selected').count(), 1);
      assert.equal(
        await page.locator('.shape-hit.selected').getAttribute('data-shape-id'),
        String(original.id),
      );
      await childRow.click({ modifiers: ['Meta'] });
      assert.equal(await page.locator('.shape-hit.selected').count(), 1);
      assert.equal(
        await page.locator('.shape-hit.selected').getAttribute('data-shape-id'),
        String(child.id),
      );
      assert.equal((await state()).revision, selectionRevision);
      await page.locator(`#selection-pane [data-selection-id="${sibling.id}"]`).click();
      await page
        .locator(`#selection-pane [data-selection-id="${child.id}"]`)
        .click({ modifiers: ['Shift'] });
      revision = (await state()).revision;
      await page.locator('[data-edit="arrange"]:visible').first().click();
      await page.getByRole('menuitem', { name: 'Align', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Align Top', exact: true }).click();
      await settled(revision);
      const aligned = (await state()).editor.slides[0].shapes[0];
      const top = (s) => {
        const [, b, , d, , f] = s.parentTransform;
        const box = s.bounds;
        return Math.min(
          ...[box.x, box.x + box.w].flatMap((x) =>
            [box.y, box.y + box.h].map((y) => b * x + d * y + f),
          ),
        );
      };
      assert.ok(Math.abs(top(aligned.children[0]) - top(aligned.children[1])) < 2);
      assert.deepEqual(aligned.bounds, original.bounds);
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], original);
      revision = (await state()).revision;
      const alignMenu = async () => {
        await page.locator('[data-edit="arrange"]:visible').first().click();
        await page.getByRole('menuitem', { name: 'Align', exact: true }).click();
      };
      await alignMenu();
      await page.getByRole('menuitemcheckbox', { name: 'Align to Slide', exact: true }).click();
      assert.equal((await state()).revision, revision);
      await alignMenu();
      assert.equal(
        await page
          .getByRole('menuitemcheckbox', { name: 'Align to Slide', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await page.getByRole('menuitem', { name: 'Align Center', exact: true }).click();
      await settled(revision);
      const slideAligned = (await state()).editor.slides[0].shapes[0];
      for (const s of slideAligned.children) {
        const [a, , c, , e] = s.parentTransform;
        const center = a * (s.bounds.x + s.bounds.w / 2) + c * (s.bounds.y + s.bounds.h / 2) + e;
        assert.ok(Math.abs(center - (await state()).editor.width / 2) < 2);
      }
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], original);
      await alignMenu();
      await page
        .getByRole('menuitemcheckbox', { name: 'Align Selected Objects', exact: true })
        .click();
      revision = (await state()).revision;
      await page.locator('#slide').focus();
      await page.keyboard.press('Meta+Alt+Shift+f');
      await settled(revision);
      assert.deepEqual(
        (await state()).editor.slides[0].shapes[0].children.map((s) => s.id),
        [child.id, sibling.id],
      );
      assert.equal(await page.locator('.shape-hit.selected').count(), 2);
      revision = (await state()).revision;
      await page.keyboard.press('Meta+Alt+g');
      await settled(revision);
      const nested = (await state()).editor.slides[0].shapes[0].children[0];
      assert.equal(nested.kind, 'group');
      assert.equal(await page.locator('.shape-hit.selected').count(), 1);
      assert.equal(
        await page.locator('.shape-hit.selected').getAttribute('data-shape-id'),
        String(nested.id),
      );
      revision = (await state()).revision;
      await page.keyboard.press('Meta+Alt+Shift+g');
      await settled(revision);
      assert.deepEqual(
        (
          await page
            .locator('.shape-hit.selected')
            .evaluateAll((nodes) => nodes.map((n) => Number(n.dataset.shapeId)))
        ).sort(),
        [child.id, sibling.id].sort(),
      );
      assert.equal((await state()).editor.slides[0].shapes.length, 1);
      assert.equal((await state()).editor.slides[0].shapes[0].children.length, 2);
      const rotatedParent = (await state()).editor.slides[0].shapes[0];
      await page.locator(`#selection-pane [data-selection-id="${rotatedParent.id}"]`).click();
      revision = (await state()).revision;
      await page.keyboard.press('Meta+Alt+Shift+g');
      await settled(revision);
      const released = (await state()).editor.slides[0].shapes;
      assert.equal(released.length, 2);
      for (const before of rotatedParent.children) {
        const after = released.find((shape) => shape.id === before.id);
        const [a, b, c, d, e, f] = before.parentTransform;
        const x = before.bounds.x + before.bounds.w / 2;
        const y = before.bounds.y + before.bounds.h / 2;
        assert.ok(Math.abs(after.bounds.x + after.bounds.w / 2 - (a * x + c * y + e)) < 2);
        assert.ok(Math.abs(after.bounds.y + after.bounds.h / 2 - (b * x + d * y + f)) < 2);
        assert.equal(after.rotation, 15);
      }
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], rotatedParent);
      await page.locator(`#selection-pane [data-selection-id="${child.id}"]`).click();
      revision = (await state()).revision;
      await page.keyboard.press('Meta+Shift+f');
      await settled(revision);
      assert.deepEqual(
        (await state()).editor.slides[0].shapes[0].children.map((s) => s.id),
        [sibling.id, child.id],
      );
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], rotatedParent);
      await page.locator(`#selection-pane [data-selection-id="${child.id}"]`).click();
      revision = (await state()).revision;
      await page.keyboard.press('Backspace');
      await settled(revision);
      const afterDelete = (await state()).editor.slides[0].shapes[0];
      assert.deepEqual(
        afterDelete.children.map((s) => s.id),
        [sibling.id],
      );
      assert.deepEqual(afterDelete.bounds, rotatedParent.bounds);
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], rotatedParent);
      await page.locator(`#selection-pane [data-selection-id="${child.id}"]`).click();
      revision = (await state()).revision;
      await page.keyboard.press('Meta+d');
      await settled(revision);
      const duplicated = (await state()).editor.slides[0].shapes;
      assert.equal(duplicated.length, 1);
      assert.equal(duplicated[0].children.length, 3);
      const copy = duplicated[0].children[2];
      assert.equal(copy.name, child.name);
      assert.equal(
        await page.locator('.shape-hit.selected').getAttribute('data-shape-id'),
        String(copy.id),
      );
      assert.notEqual(copy.id, child.id);
      assert.equal(await page.locator('.shape-hit.selected').count(), 1);
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], rotatedParent);
      await page.locator(`#selection-pane [data-selection-id="${child.id}"]`).click();
      await page.locator('#tab-animations').click();
      revision = (await state()).revision;
      await page.getByRole('button', { name: 'Entrance Appear', exact: true }).click();
      await settled(revision);
      const marker = page.locator(`.animation-marker[data-animation-shape="${child.id}"]`);
      assert.equal(await marker.count(), 1);
      assert.equal(await marker.textContent(), '1');
      const extent = rotatedParent.children.find((s) => s.id === child.id);
      const projected = [extent.bounds.x, extent.bounds.x + extent.bounds.w].flatMap((x) =>
        [extent.bounds.y, extent.bounds.y + extent.bounds.h].map((y) => ({
          x: a * x + c * y + e,
          y: b * x + d * y + f,
        })),
      );
      const currentScale =
        (await page.locator('#slide').boundingBox()).width / (await state()).editor.width;
      const markerPosition = await marker.evaluate((node) => ({
        x: parseFloat(node.style.left),
        y: parseFloat(node.style.top),
      }));
      assert.ok(
        Math.abs(markerPosition.x - (Math.min(...projected.map((p) => p.x)) * currentScale - 25)) <
          1,
      );
      assert.ok(
        Math.abs(markerPosition.y - Math.min(...projected.map((p) => p.y)) * currentScale) < 1,
      );
      await page.locator(`#selection-pane [data-selection-id="${sibling.id}"]`).click();
      await marker.click();
      assert.equal(
        await page.locator('.shape-hit.selected').getAttribute('data-shape-id'),
        String(child.id),
      );
      await page.locator('#animation-pane-toggle').click();
      assert.ok(
        (await page.locator('#animation-list [role="option"]').textContent()).includes(child.name),
      );
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.equal(await page.locator('.animation-marker').count(), 0);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], rotatedParent);
      const renameRow = page.locator(`#selection-pane [data-selection-id="${child.id}"]`);
      await page.locator('#tab-home').click();
      await page.locator('[data-edit="selection"]:visible').first().click();
      await renameRow.dblclick();
      const nameInput = page.getByRole('textbox', { name: 'Object name', exact: true });
      await nameInput.fill('Cancelled name');
      await nameInput.press('Escape');
      assert.equal(await renameRow.textContent(), child.name);
      await renameRow.press('F2');
      await nameInput.fill('Renamed child');
      revision = (await state()).revision;
      await nameInput.press('Enter');
      await settled(revision);
      assert.equal(await renameRow.textContent(), 'Renamed child');
      assert.equal((await state()).editor.slides[0].shapes[0].children[0].text, child.text);
      await page.reload();
      await page.locator('.shape-hit').first().waitFor();
      assert.equal((await state()).editor.slides[0].shapes[0].children[0].name, 'Renamed child');
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], rotatedParent);
      await page.locator('[data-edit="selection"]:visible').first().click();
      revision = (await state()).revision;
      await page.getByRole('button', { name: `Hide ${rotatedParent.name}`, exact: true }).click();
      await settled(revision);
      assert.equal((await state()).editor.slides[0].shapes[0].hidden, true);
      assert.equal(await page.locator('.shape-hit').count(), 0);
      await page.reload();
      await page.waitForFunction(
        () => window.office?.getState().editor?.slides[0].shapes[0].hidden === true,
      );
      assert.equal(await page.locator('.shape-hit').count(), 0);
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], rotatedParent);
      assert.ok((await page.locator('.shape-hit').count()) > 0);
      await page.locator('[data-edit="selection"]:visible').first().click();
      const expandGroup = page.getByRole('button', {
        name: `Expand ${rotatedParent.name}`,
        exact: true,
      });
      if (await expandGroup.count()) await expandGroup.click();
      revision = (await state()).revision;
      await page.getByRole('button', { name: `Hide ${child.name}`, exact: true }).click();
      await settled(revision);
      assert.equal((await state()).editor.slides[0].shapes[0].children[0].hidden, true);
      revision = (await state()).revision;
      await page.getByRole('button', { name: `Show ${child.name}`, exact: true }).click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], rotatedParent);
      revision = (await state()).revision;
      await page.getByRole('button', { name: `Hide ${child.name}`, exact: true }).click();
      await settled(revision);
      const beforeBulkVisibility = (await state()).editor.slides[0].shapes[0];
      revision = (await state()).revision;
      await page.getByRole('button', { name: 'Hide All', exact: true }).click();
      await settled(revision);
      const hiddenGroup = (await state()).editor.slides[0].shapes[0];
      assert.equal(hiddenGroup.hidden, true);
      assert.ok(hiddenGroup.children.every((item) => item.hidden));
      assert.equal(await page.locator('.shape-hit').count(), 0);
      revision = (await state()).revision;
      await page.getByRole('button', { name: 'Show All', exact: true }).click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], rotatedParent);
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], hiddenGroup);
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], beforeBulkVisibility);
      const dragBefore = (await state()).editor.slides[0].shapes[0];
      revision = (await state()).revision;
      await page
        .locator(`#selection-pane [data-selection-id="${child.id}"]`)
        .dragTo(page.locator(`#selection-pane [data-selection-id="${sibling.id}"]`), {
          targetPosition: { x: 10, y: 2 },
        });
      await settled(revision);
      assert.deepEqual(
        (await state()).editor.slides[0].shapes[0].children.map((item) => item.id),
        [sibling.id, child.id],
      );
      await page.reload();
      await page.locator('.shape-hit').first().waitFor();
      assert.deepEqual(
        (await state()).editor.slides[0].shapes[0].children.map((item) => item.id),
        [sibling.id, child.id],
      );
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes[0], dragBefore);
      await page.locator('#arrange-menu').click();
      await page.getByRole('menuitem', { name: 'Selection Pane...', exact: true }).click();
      await page.locator(`#selection-pane [data-selection-id="${dragBefore.id}"]`).click();
      for (const [label, change] of [
        ['Rotate Right 90°', { rotation: (dragBefore.rotation + 90) % 360 }],
        ['Rotate Left 90°', { rotation: (dragBefore.rotation + 270) % 360 }],
        [
          'Flip Horizontal',
          {
            flip: {
              horizontal: !dragBefore.flip?.horizontal,
              vertical: !!dragBefore.flip?.vertical,
            },
          },
        ],
        [
          'Flip Vertical',
          {
            flip: {
              horizontal: !!dragBefore.flip?.horizontal,
              vertical: !dragBefore.flip?.vertical,
            },
          },
        ],
      ]) {
        revision = (await state()).revision;
        await page.locator('#arrange-menu').click();
        await page.getByRole('menuitem', { name: 'Rotate or Flip' }).hover();
        await page.getByRole('menuitem', { name: label, exact: true }).click();
        await settled(revision);
        const updated = (await state()).editor.slides[0].shapes[0];
        for (const [key, value] of Object.entries(change)) assert.deepEqual(updated[key], value);
        assert.deepEqual(updated.bounds, dragBefore.bounds);
        revision = (await state()).revision;
        await page.locator('[data-edit="undo"]').click();
        await settled(revision);
        assert.deepEqual((await state()).editor.slides[0].shapes[0], dragBefore);
      }
      revision = (await state()).revision;
      await page.locator('#arrange-menu').click();
      await page.getByRole('menuitem', { name: 'Ungroup', exact: true }).click();
      await settled(revision);
      const beforeRegroup = (await state()).editor.slides[0].shapes;
      await page.reload();
      await page.locator('.shape-hit').first().waitFor();
      await page.locator('#arrange-menu').click();
      await page.getByRole('menuitem', { name: 'Selection Pane...', exact: true }).click();
      await page.locator(`#selection-pane [data-selection-id="${sibling.id}"]`).click();
      revision = (await state()).revision;
      await page.locator('#arrange-menu').click();
      await page.getByRole('menuitem', { name: 'Regroup', exact: true }).click();
      await settled(revision);
      const regrouped = (await state()).editor.slides[0].shapes;
      assert.equal(regrouped.length, 1);
      assert.deepEqual(
        regrouped[0].children.map((item) => item.id),
        beforeRegroup.map((item) => item.id),
      );
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes, beforeRegroup);
      await page.locator(`#selection-pane [data-selection-id="${sibling.id}"]`).click();
      await page.locator('#arrange-menu').click();
      assert.equal(
        await page
          .getByRole('menuitem', { name: 'Regroup', exact: true })
          .locator('.menu-shortcut')
          .textContent(),
        '⌥⌘J',
      );
      await page.keyboard.press('Escape');
      revision = (await state()).revision;
      await page.keyboard.press('Meta+Alt+KeyJ');
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes, regrouped);
      revision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await settled(revision);
      assert.deepEqual((await state()).editor.slides[0].shapes, beforeRegroup);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      proc.kill('SIGTERM');
      await new Promise((resolve) =>
        proc.exitCode !== null ? resolve() : proc.once('exit', resolve),
      );
      await rm(dir, { recursive: true, force: true });
    }
  },
);
