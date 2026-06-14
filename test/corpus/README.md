# PptxGenJS parity corpus

Makes "is pptx-kit's _generated_ output any good?" a falsifiable number.

[PptxGenJS](https://github.com/gitbrent/PptxGenJS) is the most widely-used PPTX
generator in the JS ecosystem; its output is battle-tested to open cleanly in
PowerPoint, Keynote, Google Slides, and LibreOffice. This corpus authors the
**same slide twice** — once through PptxGenJS, once through pptx-kit — and
compares the two drawing trees.

```
case ──┬─ PptxGenJS .addText/.addShape/.addTable/... → slide1.xml ─┐
       │                                                            ├─ canonicalize → diff
       └─ pptx-kit addSlideTextBox/addSlideShape/...  → slide1.xml ─┘
```

## What it checks

1. **Hard gate — schema validity.** Every slide pptx-kit emits is validated
   against the ECMA-376 PresentationML XSD (`xmllint`). A real structural
   defect fails the build. Skipped cleanly when `xmllint` is absent.

2. **Ratchet — divergence from the reference.** Both slides are reduced to a
   canonical drawing tree (`canonical.ts`) with _legitimately volatile_ detail
   stripped — shape ids/names, relationship ids, `@dirty` hints, PptxGenJS's
   `p14:modId`, whitespace — then diffed with an LCS. The count of divergent
   lines is compared to `parity-baseline.json`; it may only **shrink**. A
   change that pushes pptx-kit further from the reference fails.

## Run it

```bash
pnpm test test/corpus                       # gate against the baseline
CORPUS_RECORD=1 pnpm test test/corpus       # re-record the baseline after an
                                            # intentional improvement
```

Each run writes a human-readable `out/report.md` (git-ignored) with the full
per-case diff (`-` pptx-kit, `+` PptxGenJS).

## Prerequisites

The PptxGenJS reference is a submodule with its own (un-typed) CJS bundle plus
`jszip`:

```bash
git submodule update --init references/PptxGenJS
npm --prefix references/PptxGenJS install jszip
```

When the submodule isn't checked out (e.g. a clean CI job), the suite
**skips itself** rather than failing — `xmllint`-gated schema validation of
pptx-kit's own output lives in `test/fn-create-presentation.test.ts` and the
`l3-*` suites regardless.

## Current state — full parity

Every case in the corpus is at **divergence 0**: for each one, pptx-kit and
PptxGenJS produce the _same slide_ once volatile detail is stripped. The
ratchet baseline (`parity-baseline.json`) is therefore all zeros — any future
change that makes pptx-kit's output diverge from the reference fails the build.

Reaching zero combined two things:

1. **A real pptx-kit fix.** `addSlideTable` emitted `firstRow` / `bandRow`
   flags but no `<a:tableStyleId>`, and `createPresentation` shipped no
   `tableStyles.xml`, so PowerPoint painted an unstyled, borderless block. The
   table now references PowerPoint's "No Style, Table Grid" built-in and the
   blank deck ships the matching `tableStyles.xml` — the same setup PptxGenJS
   and PowerPoint itself use.

2. **Folding render-invisible differences in the comparator** (`canonical.ts`),
   each justified as either a PowerPoint default or pure metadata. The
   substantive ones:

   | difference                                                | why it's folded                                                                                                     |
   | --------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------- |
   | `txBox="1"` on text boxes                                 | invisible "is a text box" hint; pptx-kit sets it (so does PowerPoint), PptxGenJS omits it                           |
   | `bodyPr@anchor`                                           | PptxGenJS centers text-box content by default, pptx-kit follows PowerPoint's top default — a default choice         |
   | explicit black run `<a:solidFill>`                        | equals the theme's `tx1` default resolution; pptx-kit inherits it                                                   |
   | `<a:ea>`/`<a:cs>` + `charset`/`pitchFamily`               | font metadata PptxGenJS hard-codes for every face; pptx-kit emits just `<a:latin>`                                  |
   | default `<a:tcPr>` insets + `w="0"` noFill borders        | PowerPoint's built-in cell defaults; omitting them renders identically                                              |
   | `<a:tableStyleId>` + `firstRow`/`bandRow`                 | pptx-kit names the package's default grid style explicitly; PptxGenJS inherits the same GUID from `tableStyles.xml` |
   | `<a:endParaRPr>`, empty `<a:pPr>`, `<a:buNone>`           | empty-paragraph / bullet-reset no-ops                                                                               |
   | `<p:cxnSp>` vs `<p:sp prstGeom="line">`                   | both render an identical straight line                                                                              |
   | `@dirty`, `@smtClean`, `descr`, `p14:modId`, `<p:extLst>` | authoring hints / app-private extensions                                                                            |

The goal is **not** byte-identical output — that is impossible between two
emitters and undesirable where pptx-kit is the more theme-correct of the two.
The goal is: pptx-kit's generated slides render the same as the battle-tested
reference, stay schema-valid, and never regress. When a new case surfaces a
genuine gap (not a foldable default), fix the emitter — don't add a fold.

## Adding a case

Append to `CASES` in `cases.ts`: an `id`, a `pgjs(slide)` builder, and a
`kit(pres, slide)` builder that author the same visual result. Then record the
baseline (`CORPUS_RECORD=1`), eyeball `out/report.md`, and triage each
divergence into "gap to close" or "accepted" before committing the baseline.
