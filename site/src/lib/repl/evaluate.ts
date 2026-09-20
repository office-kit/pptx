// How the REPL turns editor text into running code. Both modes end up in
// `new Function`, wrapped the same way, so an error's line maps back to the
// editor with one shared offset.

const JSX_IMPORT_SOURCE = '@office-kit/pptx-dsl';

// `new Function` puts a two-line header in front of the body, and `asyncBody`
// adds two more lines before the user's code.
export const WRAPPER_LINES = 4;

/** Wraps code in an async function so `await` is allowed at the top level. */
export const asyncBody = (code: string): string =>
  `'use strict';\nreturn (async () => {\n${code}\n})();`;

export type ModuleMap = Readonly<Record<string, unknown>>;

/**
 * Runs `source` as a TSX module, the way `pptx-dev` does on disk, and returns
 * its default export. Sucrase turns the module into CommonJS (keeping every
 * token on its original line) and `require` resolves against `modules` — the
 * caller owns that map so the test suite can run the same code against the
 * library source instead of the built packages.
 *
 * Throws when the code does not parse, imports a module outside `modules`, or
 * has no default export.
 */
export const evaluateTsx = async (source: string, modules: ModuleMap): Promise<unknown> => {
  // Sucrase is only needed once someone opens the TSX tab.
  const { transform } = await import('sucrase');
  const { code } = transform(source, {
    transforms: ['typescript', 'jsx', 'imports'],
    jsxRuntime: 'automatic',
    jsxImportSource: JSX_IMPORT_SOURCE,
    production: true,
  });
  const require = (name: string): unknown => {
    if (!Object.hasOwn(modules, name)) {
      const available = Object.keys(modules).join(', ');
      throw new Error(`Cannot import "${name}". The REPL provides: ${available}.`);
    }
    return modules[name];
  };
  const exports: { default?: unknown } = {};
  await new Function('require', 'exports', asyncBody(code))(require, exports);
  if (exports.default === undefined) {
    throw new Error('The TSX file must default-export a Presentation.');
  }
  return exports.default;
};
