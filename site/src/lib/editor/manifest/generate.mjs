// Capability-manifest generator — the backbone of the coverage guarantee.
//
// Reads the @office-kit/pptx public API source, enumerates every *mutating*
// export (the verbs a PowerPoint-style UI must expose as an operation), and
// emits `capabilities.generated.json`: one entry per capability with its
// operand, parsed parameter list, and a heuristic category.
//
// This file is the *source of truth for what exists*. The coverage test
// (`coverage.test.ts`) independently re-derives the mutating-export set from
// the compiled library and fails if it drifts from the manifest — so a new
// authoring function added to the library cannot be silently left out of the
// editor. Human-authored refinements (labels, ribbon groups, richer param
// schemas) live in `overrides.ts` and are merged on top; they never remove
// entries.
//
// Run: `node site/src/lib/editor/manifest/generate.mjs`
import { readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(here, '..', '..', '..', '..', '..');
const srcRoot = join(repoRoot, 'src');
const apiIndex = join(srcRoot, 'api', 'index.ts');

// Verb prefixes that denote a state-changing (authoring) operation. Kept in
// sync with `coverage.test.ts` — the two MUST agree.
export const MUTATING_VERBS = [
  'add',
  'set',
  'clear',
  'replace',
  'remove',
  'insert',
  'duplicate',
  'bring',
  'send',
  'append',
  'group',
  'ungroup',
  'swap',
  'sort',
  'reverse',
  'rename',
  'move',
  'merge',
  'import',
  'copy',
  'create',
  'translate',
  'touch',
  'increment',
  'compact',
];

// A handful of mutating-verb exports are not slide-authoring operations the
// canvas UI drives; they still get a command (reachable via palette) but we
// tag them so the ribbon layer can skip them.
const NON_CANVAS = new Set([
  'compactPackage',
  'incrementRevision',
  'touchModified',
  'mergePresentations',
  'importSlide',
  'setCoreProperties',
  'setExtendedProperties',
  'setMediaPartBytes',
  'removeThumbnail',
  'setThumbnail',
]);

function isMutating(name) {
  return MUTATING_VERBS.some(
    (v) =>
      name.startsWith(v) &&
      name.length > v.length &&
      name[v.length] === name[v.length].toUpperCase(),
  );
}

function walk(dir) {
  let out = [];
  for (const e of readdirSync(dir)) {
    const p = join(dir, e);
    const s = statSync(p);
    if (s.isDirectory()) out = out.concat(walk(p));
    else if (p.endsWith('.ts') && !p.endsWith('.test.ts')) out.push(p);
  }
  return out;
}

// The value exports of the public API — parse the `export { ... }` blocks.
function publicValueExports() {
  const code = readFileSync(apiIndex, 'utf8');
  const names = new Set();
  const re = /export \{([^}]*)\}/g;
  let m;
  while ((m = re.exec(code))) {
    for (const raw of m[1].split(',')) {
      const t = raw.trim().replace(/^type\s+/, '');
      if (!t || t.startsWith('type ')) continue;
      // skip `type X` entries and aliases (`a as b` → keep b)
      const asMatch = t.match(/\bas\s+([A-Za-z_]\w*)/);
      const name = asMatch ? asMatch[1] : t;
      if (/^[a-z][A-Za-z0-9_]*$/.test(name)) names.add(name);
    }
  }
  return names;
}

const files = walk(srcRoot);
const source = Object.fromEntries(files.map((f) => [f, readFileSync(f, 'utf8')]));

// Extract the full parameter block + return type for a named arrow-const.
function extractSignature(name) {
  for (const [file, code] of Object.entries(source)) {
    const idx = code.indexOf(`export const ${name} = `);
    if (idx < 0) continue;
    let i = idx + `export const ${name} = `.length;
    // optional `async`
    if (code.startsWith('async ', i)) i += 6;
    if (code[i] !== '(') continue;
    let depth = 0;
    const parenStart = i;
    for (; i < code.length; i++) {
      if (code[i] === '(') depth++;
      else if (code[i] === ')') {
        depth--;
        if (depth === 0) {
          i++;
          break;
        }
      }
    }
    const params = code.slice(parenStart + 1, i - 1);
    // return type: from `:` to `=>`
    const arrow = code.indexOf('=>', i);
    const ret = code
      .slice(i, arrow)
      .replace(/^\s*:\s*/, '')
      .trim();
    return { file: file.replace(srcRoot + '/', '').replace(/\\/g, '/'), params, ret };
  }
  return null;
}

// Split a parameter list on top-level commas (ignoring nested <>, (), {}, []).
function splitTop(s) {
  const out = [];
  let depth = 0,
    cur = '';
  for (const ch of s) {
    if ('<([{'.includes(ch)) depth++;
    else if ('>)]}'.includes(ch)) depth--;
    if (ch === ',' && depth === 0) {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  if (cur.trim()) out.push(cur);
  return out.map((x) => x.trim()).filter(Boolean);
}

// Classify a TS type string into a UI param kind.
function kindOf(typeText, paramName) {
  const t = typeText.replace(/\s+/g, ' ').trim();
  const lname = (paramName || '').toLowerCase();
  if (/^Emu\b/.test(t) || /\bEmu$/.test(t)) return 'emu';
  if (lname === 'color' || /\bcolor\b/i.test(lname)) return 'color';
  if (/^number\b/.test(t)) return 'number';
  if (/^boolean\b/.test(t)) return 'boolean';
  if (/^string\b/.test(t)) return 'string';
  // string-literal union → enum
  const lits = t.match(/'[^']+'/g);
  if (
    lits &&
    new RegExp(`^(?:'[^']+'\\s*\\|?\\s*)+$`).test(t.replace(/\bnull\b|\bundefined\b/g, '').trim())
  ) {
    return 'enum';
  }
  if (/^\{/.test(t) || /Options$|Spec$|Data$|Input$/.test(t)) return 'object';
  return 'object';
}

function parseParam(p) {
  // `name?: Type` or `name: Type = default`
  const eq = (() => {
    // find top-level ` = ` for defaults
    let depth = 0;
    for (let i = 0; i < p.length - 2; i++) {
      const ch = p[i];
      if ('<([{'.includes(ch)) depth++;
      else if ('>)]}'.includes(ch)) depth--;
      if (depth === 0 && p.slice(i, i + 3) === ' = ') return i;
    }
    return -1;
  })();
  let deflt;
  let body = p;
  if (eq >= 0) {
    deflt = p.slice(eq + 3).trim();
    body = p.slice(0, eq).trim();
  }
  const colon = body.indexOf(':');
  if (colon < 0)
    return {
      name: body.replace(/[?]/g, '').trim(),
      type: 'unknown',
      kind: 'object',
      optional: body.includes('?'),
    };
  const rawName = body.slice(0, colon).trim();
  const optional = rawName.endsWith('?');
  const name = rawName.replace(/\?$/, '').replace(/^\{[\s\S]*$/, 'options');
  const type = body.slice(colon + 1).trim();
  const kind = kindOf(type, name);
  const spec = { name, type: type.replace(/\s+/g, ' '), kind, optional };
  if (deflt !== undefined) spec.default = deflt;
  if (kind === 'enum') {
    spec.enumValues = (type.match(/'([^']+)'/g) || []).map((x) => x.slice(1, -1));
  }
  return spec;
}

const OPERAND_BY_TYPE = {
  PresentationData: 'presentation',
  SlideData: 'slide',
  SlideShapeData: 'shape',
  TableCellData: 'cell',
};

// Heuristic category from the source file name + export name.
function categoryOf(name, file) {
  const f = file;
  if (/charts\.ts$/.test(f)) return 'chart';
  if (/tables\.ts$/.test(f)) return 'table';
  if (/transition/.test(f)) return 'transition';
  if (/animation/.test(f)) return 'animation';
  if (/comments\.ts$/.test(f)) return 'comment';
  if (/notes/.test(f)) return 'notes';
  if (/theme|color-map|features/.test(f)) return 'theme';
  if (/background/.test(f)) return 'slide-background';
  if (/sections/.test(f)) return 'section';
  if (
    /\bslide-(deck|query|size|title)\b/.test(f) ||
    (/Slide/.test(name) && /slides?\b/i.test(name))
  )
    return 'slide';
  if (/hyperlink/i.test(name)) return 'hyperlink';
  if (/image/i.test(name)) return 'image';
  if (/gradient|patternfill|nofill|fill/i.test(name)) return 'fill';
  if (/stroke|arrow|dash|cap|join|compound/i.test(name)) return 'stroke';
  if (/glow|shadow|effect|reflection/i.test(name)) return 'effect';
  if (/paragraph|bullet/i.test(name)) return 'paragraph';
  if (/run|text|font|anchor|autofit|wrap|margin|column|direction/i.test(name)) return 'text';
  if (/^set?Shape|Shape/.test(name)) return 'shape';
  if (/Slide/.test(name)) return 'slide';
  if (/Presentation|Core|Extended|Revision|Modified|Media|Thumbnail|Package/.test(name))
    return 'presentation';
  return 'misc';
}

const publicNames = [...publicValueExports()].filter(isMutating).sort();
const capabilities = [];
const missing = [];
for (const name of publicNames) {
  const sig = extractSignature(name);
  if (!sig) {
    missing.push(name);
    continue;
  }
  const allParams = splitTop(sig.params).map(parseParam);
  // The operand is the leading parameter *only* when its type is one of the
  // domain objects (PresentationData / SlideData / SlideShapeData /
  // TableCellData). Factories like `createPresentation(options)` take no
  // operand — their first parameter is a real user argument.
  const firstType = allParams.length ? (allParams[0].type || '').split(/[<\s]/)[0] : '';
  const operandFromFirst = OPERAND_BY_TYPE[firstType];
  const takesOperand = Boolean(operandFromFirst);
  const operand = operandFromFirst ?? 'presentation';
  // `params` is the user-facing argument list — the operand is dropped so the
  // registry/forms never ask the user to supply the object they already have
  // selected.
  const params = takesOperand ? allParams.slice(1) : allParams;
  capabilities.push({
    id: name,
    operand,
    takesOperand,
    category: categoryOf(name, sig.file),
    file: sig.file,
    returns: sig.ret,
    canvas: !NON_CANVAS.has(name),
    params,
  });
}

if (missing.length) {
  console.error('WARNING: could not extract signatures for:', missing.join(', '));
}

const outPath = join(here, 'capabilities.generated.json');
writeFileSync(
  outPath,
  JSON.stringify(
    { generatedFrom: 'src/api/index.ts', count: capabilities.length, capabilities },
    null,
    2,
  ) + '\n',
);
console.log(`Wrote ${capabilities.length} capabilities to ${outPath.replace(repoRoot + '/', '')}`);
