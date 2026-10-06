import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { tokenStyles } from '../src/styles.ts';

const declarations = (block) =>
  Object.fromEntries(
    [...block.matchAll(/(--ok-[\w-]+)\s*:\s*([^;]+);/g)].map(([, name, value]) => [
      name,
      value.replace(/\s+/g, ' ').trim(),
    ]),
  );

test('the preview shell copies the editor design tokens for both color schemes', async () => {
  const css = await readFile(
    new URL('../../../site/src/lib/editor/ui/tokens.css', import.meta.url),
    'utf8',
  );
  const [light, dark] = css.split('@media (prefers-color-scheme: dark)');
  const [shellLight, shellDark] = tokenStyles.split('@media (prefers-color-scheme: dark)');
  const editorLight = declarations(light.slice(0, light.indexOf('}')));
  const editorDark = declarations(dark.slice(0, dark.indexOf('}\n}')));
  const copiedLight = declarations(shellLight);
  const copiedDark = declarations(shellDark);
  // Layout-only tokens (ribbon height, pane widths) do not apply to the shell.
  for (const name of ['--ok-ribbon-h', '--ok-nav-w', '--ok-panel-w']) delete editorLight[name];
  assert.deepEqual(copiedLight, editorLight);
  assert.deepEqual(copiedDark, editorDark);
});
