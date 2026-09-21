import ts from 'typescript';
import { readFile, writeFile, realpath } from 'node:fs/promises';
import { dirname, relative, isAbsolute, extname } from 'node:path';

interface Preview {
  revision: number;
  slides: string[];
  slideTexts: string[];
  dependencies: string[];
}
interface Change {
  file: string;
  before: string;
  after: string;
}

/** Literal edits are deliberately conservative: computed and ambiguous text goes to the agent. */
export function createTextEditor(
  entry: string,
  preview: () => Preview,
  verify: () => Promise<string | null>,
) {
  let busy = false;
  let undo: Change | undefined;
  async function restore(change: Change) {
    if (
      (await realpath(change.file)) !== change.file ||
      (await readFile(change.file, 'utf8')) !== change.after
    )
      throw new Error('Source changed during editing. Your newer changes were preserved.');
    await writeFile(change.file, change.before);
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
          !value.before.trim() ||
          value.before.length > 12000 ||
          typeof value.after !== 'string' ||
          value.after.length > 12000 ||
          value.after === value.before
        )
          throw new Error('Invalid text edit.');
        if (current.slideTexts[value.slide]!.split(value.before).length !== 2)
          throw new Error('Text cannot be located uniquely on this slide. Use Apply with AI.');
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
              return { file: path, before: await readFile(path, 'utf8') };
            }),
        );
        const candidates: Change[] = [];
        for (const item of files) {
          if (!item) continue;
          const source = ts.createSourceFile(
            item.file,
            item.before,
            ts.ScriptTarget.Latest,
            true,
            /[jt]sx$/.test(extname(item.file)) ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
          );
          const visit = (node: ts.Node) => {
            const text = ts.isJsxText(node)
              ? node.text
                  .split(/\r\n|\n|\r/)
                  .map((line, index, lines) => {
                    let value = line.replaceAll('\t', ' ');
                    if (index) value = value.trimStart();
                    if (index < lines.length - 1) value = value.trimEnd();
                    return value;
                  })
                  .filter(Boolean)
                  .join(' ')
              : ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node)
                ? node.text
                : undefined;
            if (text === value.before) {
              // Attribute strings use JSX entity rules, not JavaScript string escapes.
              const jsx = ts.isJsxText(node) || ts.isJsxAttribute(node.parent);
              const replacement = jsx
                ? '{' + JSON.stringify(value.after) + '}'
                : JSON.stringify(value.after);
              const start = ts.isJsxText(node) ? node.pos : node.getStart(source);
              candidates.push({
                ...item,
                after: item.before.slice(0, start) + replacement + item.before.slice(node.end),
              });
            }
            ts.forEachChild(node, visit);
          };
          visit(source);
        }
        if (candidates.length !== 1)
          throw new Error('Text is computed or has multiple source matches. Use Apply with AI.');
        const change = candidates[0]!;
        if (
          preview().revision !== current.revision ||
          (await realpath(change.file)) !== change.file ||
          (await readFile(change.file, 'utf8')) !== change.before
        )
          throw new Error('Source changed. Select the text again.');
        await writeFile(change.file, change.after);
        try {
          const error = await verify();
          if (error) throw new Error('The text change did not build.');
          const next = preview();
          if (
            next.slides.length !== current.slides.length ||
            next.slides.some(
              (svg, index) => index !== value.slide && svg !== current.slides[index],
            ) ||
            next.slideTexts[value.slide] !==
              current.slideTexts[value.slide]!.replace(value.before, value.after)
          )
            throw new Error(
              'This source affects other text or slides. Use Apply with AI for a local edit.',
            );
          if (
            (await realpath(change.file)) !== change.file ||
            (await readFile(change.file, 'utf8')) !== change.after
          )
            throw new Error('Source changed during editing.');
          undo = change;
        } catch (error) {
          await restore(change);
          await verify();
          throw error;
        }
      } finally {
        busy = false;
      }
    },
    async undo() {
      if (busy) throw new Error('A text edit is already running.');
      if (!undo) throw new Error('No direct text edit to undo.');
      busy = true;
      try {
        await restore(undo);
        undo = undefined;
        const error = await verify();
        if (error) throw new Error('Text restored, but the project has build errors.');
      } finally {
        busy = false;
      }
    },
  };
}
