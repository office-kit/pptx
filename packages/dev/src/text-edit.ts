import { planTextEdit, verifyTextEdit, type TextEditPlan } from '@office-kit/pptx-dsl/source-edit';
import { readFile, writeFile, realpath } from 'node:fs/promises';
import { dirname, relative, isAbsolute } from 'node:path';

interface Preview {
  revision: number;
  slides: string[];
  slideTexts: string[];
  dependencies: string[];
}
type Change = Extract<TextEditPlan, { ok: true }>['change'];

const planErrors: Record<Extract<TextEditPlan, { ok: false }>['reason'], string> = {
  invalid: 'Invalid text edit.',
  'not-unique-on-slide': 'Text cannot be located uniquely on this slide. Use Apply with AI.',
  'not-found': 'Text is computed or has multiple source matches. Use Apply with AI.',
  ambiguous: 'Text is computed or has multiple source matches. Use Apply with AI.',
};

const message = (error: unknown) => (error instanceof Error ? error.message : String(error));

/** Literal edits are deliberately conservative: computed and ambiguous text goes to the agent. */
export function createTextEditor(
  entry: string,
  preview: () => Preview,
  verify: () => Promise<string | null>,
) {
  let busy = false;
  async function restore(change: Change) {
    if (
      (await realpath(change.path)) !== change.path ||
      (await readFile(change.path, 'utf8')) !== change.after
    )
      throw new Error('Source changed during editing. Your newer changes were preserved.');
    await writeFile(change.path, change.before);
  }
  return {
    async edit(value: { revision: number; slide: number; before: string; after: string }) {
      if (busy) throw new Error('A text edit is already running.');
      busy = true;
      try {
        const current = preview();
        if (value.revision !== current.revision)
          throw new Error('Preview changed. Select the text again.');
        if (
          !Number.isInteger(value.slide) ||
          !current.slides[value.slide] ||
          typeof value.before !== 'string' ||
          typeof value.after !== 'string'
        )
          throw new Error(planErrors.invalid);
        const root = await realpath(dirname(entry));
        const files = await Promise.all(
          [...new Set(current.dependencies)]
            .filter((file) => /\.[cm]?[jt]sx?$/.test(file))
            .map(async (file) => {
              const path = await realpath(file),
                local = relative(root, path);
              if (
                local.startsWith('..') ||
                isAbsolute(local) ||
                local.split(/[\\/]/).includes('node_modules')
              )
                return null;
              return { path, source: await readFile(path, 'utf8') };
            }),
        );
        const plan = planTextEdit({
          files: files.filter((file) => file !== null),
          slideText: current.slideTexts[value.slide]!,
          before: value.before,
          after: value.after,
        });
        if (!plan.ok) throw new Error(planErrors[plan.reason]);
        const { change } = plan;
        if (
          preview().revision !== current.revision ||
          (await realpath(change.path)) !== change.path ||
          (await readFile(change.path, 'utf8')) !== change.before
        )
          throw new Error('Source changed. Select the text again.');
        await writeFile(change.path, change.after);
        try {
          const error = await verify();
          if (error) throw new Error('The text change did not build.');
          if (
            !verifyTextEdit({
              previous: current,
              next: preview(),
              slide: value.slide,
              before: value.before,
              after: value.after,
            }).ok
          )
            throw new Error(
              'This source affects other text or slides. Use Apply with AI for a local edit.',
            );
          if (
            (await realpath(change.path)) !== change.path ||
            (await readFile(change.path, 'utf8')) !== change.after
          )
            throw new Error('Source changed during editing.');
        } catch (error) {
          try {
            await restore(change);
            await verify();
          } catch (cleanup) {
            // Keep the reason the edit was refused; the cleanup failure alone
            // would hide it.
            throw new Error(`${message(error)} ${message(cleanup)}`, { cause: error });
          }
          throw error;
        }
      } finally {
        busy = false;
      }
    },
  };
}
