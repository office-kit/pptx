# CLAUDE.md

## Project overview

**pptx-kit** is a TypeScript library that generates `.pptx` (PowerPoint /
Office Open XML Presentation) files from a typed object model. It runs in
**both Node.js and the browser** from a single ESM bundle, and produces files
that open cleanly in PowerPoint, Keynote, Google Slides, and LibreOffice
Impress.

The library has two complementary entry paths:

1. **From scratch** — build a presentation programmatically and emit a `.pptx`.
2. **From a template** — load an existing `.pptx`, mutate slides / placeholders
   / shapes, and emit the result.

The long-term goal is to support the full **OOXML PresentationML** spec
([ECMA-376 Part 1, §19](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/))
plus the supporting DrawingML / SpreadsheetML pieces a presentation actually
references (charts, embedded media metadata, transitions, animations, theme,
slide masters / layouts, comments, notes).

## Tech stack

- **Language**: TypeScript (strict mode), targeting ES2022.
- **Runtimes**: Node.js >= 20 and modern browsers (Chrome / Firefox / Safari
  current-2). One ESM bundle, no Node-only built-ins on the hot path — use
  Web-standard `Uint8Array`, `TextEncoder/Decoder`, `CompressionStream` where
  available.
- **Package manager**: pnpm.
- **Build**: `tsdown` → ESM + `.d.ts`.
- **Test**: `vitest` (unit + golden-file fixture comparisons against real
  PPTX bytes).
- **Lint / format**: oxlint + oxfmt.
- **Release**: changesets + npm OIDC trusted publisher (see
  `.github/workflows/release.yml`).
- **Dependencies**: keep the runtime dep list **tiny**. ZIP and XML are the
  only two real third-party needs; everything else should be hand-rolled or
  not done at all. Every new runtime dependency must be justified in the PR.

## Domain-specific rules

### OOXML is the source of truth, not our types

The ECMA-376 spec defines the wire format. Our TypeScript model is a
convenience layer over it.

- When the spec and our types disagree, **the spec wins**. Fix the types.
- When PowerPoint and the spec disagree (and they do, often), **PowerPoint
  wins** — but leave a comment naming the specific behavior, because the next
  reader will not know.
- Round-trip safety matters: `parse(serialize(x))` must be structurally equal
  to `x` for everything we claim to support. Add a fixture test the first time
  a round-trip case comes up; do not rely on visual inspection.

### Output must be valid PPTX, not "valid enough"

A `.pptx` that opens in PowerPoint but crashes Keynote is a bug. A `.pptx`
that opens everywhere but is rejected by Microsoft's
[Open XML SDK Productivity Tool](https://github.com/dotnet/Open-XML-SDK) is
also a bug.

- All XML output must be schema-valid against the official XSDs.
- Required relationships (`_rels/*.rels`, `[Content_Types].xml`) are not
  optional. There is no shortcut where you "just skip the rels."
- IDs (`rId*`, shape IDs, drawing IDs) must be **unique and stable** within
  their scope. Generate them through the central ID allocator; never
  hand-roll a string concatenation.

### Cross-runtime correctness

This library runs in Node and the browser. That constrains what you can
import.

- No `fs`, `path`, `buffer`, `stream`, `zlib` on the hot path. Use Web
  primitives (`Uint8Array`, `Blob`, `CompressionStream`).
- If a feature genuinely needs the Node `fs` (e.g., reading a template from
  disk), put it behind a `node:` subpath export so the browser bundle does
  not pull it in.
- Do not assume a `Buffer` global. Use `Uint8Array` and convert at the
  boundary if a Node caller passes a `Buffer`.

### Public API is small and explicit

The library will be **the** way to author / edit PPTX in the JS ecosystem,
so the public surface needs to be defended.

- One canonical path per capability (see "One way to do one thing" below).
- Builder functions return new objects; mutation is reserved for the
  template-editing path, and it is explicit (`slide.addShape(...)`).
- All inputs are typed by discriminated unions when the OOXML schema has a
  choice (`a:srgbClr` vs `a:schemeClr` etc.). No "color is a string and we
  parse it." Parse, don't validate (see below).
- `internal/` is internal. If users import from it, the public API failed —
  fix the public API, do not stabilize the internal symbol.

## Operating context

- **Audience**: OSS — code that external contributors and end users will read.
- **Language**: TypeScript. Principles below are language-agnostic, but the
  examples in this repo are TS.
- **Readers**: yourself in 6 months, a first-time contributor, an AI agent.

OSS is not the same as "code only I touch." Optimize for **a stranger not getting
confused**, not for your own convenience.

## Core principles

Every line of code must be justified.

| Principle               | Meaning                                                                                 |
| ----------------------- | --------------------------------------------------------------------------------------- |
| Simplicity              | No premature abstraction, no features for hypothetical futures. YAGNI.                  |
| Consistency             | Match existing patterns. New patterns require an explicit reason.                       |
| Performance             | Don't write N+1 / O(n²) in the first place. Defend with code shape, not profilers.      |
| Security                | Validate at boundaries (external input, file I/O, external APIs). Watch OWASP.          |
| Maintainability         | Write code that you in 6 months and a new contributor can read and understand.          |
| Backwards compatibility | Public API is preserved unless a breaking change is genuinely justified. Follow SemVer. |

## "One way to do one thing"

> A capability has exactly one canonical path through the public API.

Do not add a **parallel API** — a second path that produces the same result as an
existing one — just because it's shorter, more discoverable, or "feels nicer." Reasons:

1. Every reader pays the "which one should I use?" cost on every code review and onboarding.
2. Two APIs = two of everything: docs, tests, types, bundle size, compatibility surface.
3. If the README says "use X for Y" but the library also accepts Z, the docs are lying.

**The rule bends** in two cases:

- An existing path is **misleadingly named** — rename it, don't parallelize. (Pre-1.0:
  rename outright. Post-1.0: deprecate + remove on next major.)
- A capability is reachable but only at an abstraction so low that every real user
  re-implements the same wrapper — graduate the wrapper into the library and **hide
  the low-level path** behind an "escape hatch" subpath.

In both cases the result is still **one path per capability** for the typical user.

Decision flow when evaluating a feature request:

1. Is the capability **already reachable** through the public API? → Yes: reject.
2. Is the unreachable capability **inside this project's scope**? → No: point to the
   appropriate other library.
3. If it **replaces** an existing path, is the replacement **strictly better** on
   every dimension that matters (ergonomics, types, performance, size)? → If only some
   dimensions are better, you're proposing a parallel API → reject.

## Defensive programming

**Yes at boundaries; no internally.**

| Situation                                    | Defend? | Example                                               |
| -------------------------------------------- | ------- | ----------------------------------------------------- |
| External input (HTTP, CLI args, files)       | **Yes** | Schema-validate, fail loudly on invalid input         |
| External API / third-party calls             | **Yes** | Error handling, retries, timeouts                     |
| Untrusted data (user uploads, etc.)          | **Yes** | Validate / sanitize                                   |
| Already validated upstream                   | No      | Don't re-validate — that's noise                      |
| Already guaranteed by the type system        | No      | Don't write redundant null checks / optional chaining |
| Cases that are impossible by type definition | No      | No "just in case" guards                              |

"Just in case" checks pile up and **bury the real validation** under noise. If a value
has already been validated upstream or is guaranteed by the type system, do not
re-check it.

**Don't swallow exceptions.** Catching and silently returning `null` for unexpected
errors is anti-defensive — it hides bugs and security issues.

- **Internal logic threw** → that's a bug. Let it propagate so it surfaces fast.
- **External input was invalid** → convert it into the appropriate error type (4xx HTTP,
  CLI exit code, etc.) and return that. Don't pretend the input was valid.

## Hard "no"s

- **N+1 access**: sequential `await` inside a loop (DB / API / file). Bulk it.
- **O(n²) operations**: `find` / `filter` / linear search inside a loop. Build a Map / Set first.
- **Type-cast escape hatches**: `as unknown as T` and equivalents. Use validation to convert safely.
- **Redundant "just-in-case" checks**: re-validating values whose contract is already met.
- **Magic numbers / strings**: name them as constants.
- **Dead code**: "might use this later," commented-out implementations. Delete.
- **Swallowed exceptions**: `catch` blocks that silently return null. Let errors propagate.
- **Silent breaking changes to public API**: SemVer violation. Always declare in CHANGELOG / changeset.

## Comments

Comments explain **why**, not what.

```ts
// ❌ Self-evident "what" — drop it.
/** Get a user. */
const getUser = (id: string) => {
  /* ... */
};

// ✅ "Why" — the code can't say this on its own.
// Salesforce caps batch upsert at 100; we chunk to stay under the limit.
const batches = chunk(records, 100);

// ❌ Comments about the current task — that belongs in the PR description.
// Added for issue #123
// TODO: also handle X later

// ✅ Constraints / pitfalls a future reader needs to know.
// Order-by depends on the (created_at, id) composite index. Reordering the
// columns turns this into a full scan.
```

**Don't write a comment when**:

- The identifier name already says it
- It references the current task / PR / issue (Git history covers that, and these rot)
- It marks deleted code (`// removed: ...`)

**Do write a comment when**:

- The implementation looks weird and the reason is non-obvious (past bug, performance
  workaround, third-party API quirk)
- There's an implicit constraint or invariant
- A future reader is likely to step on a specific rake

## Use the type system

| Pattern               | What it buys                                                               |
| --------------------- | -------------------------------------------------------------------------- |
| Branded / newtype IDs | Don't mix up `UserId` and `OrderId`                                        |
| Discriminated unions  | Eliminate "both null" / "both set" invalid states at the type level        |
| Exhaustiveness checks | `switch` / `match` with `never` so adding a new variant fails to compile   |
| Parse, don't validate | Convert input into a validated type once at the boundary; trust internally |

These translate to languages without first-class type systems too: the principle is
"inside the system, values are already in the right shape — guarantee that at the boundary."

## OSS-specific discipline

### Public API is a contract

Anything mentioned in `README` / `docs` is a contract. Users depend on the **shape**,
**name**, **behavior**, and **exceptions** it produces.

- **Adding** a new export: minor bump, fine.
- **Changing existing behavior**: breaking change → major bump.
- **Removing or renaming**: breaking change → major bump (with prior deprecation if
  the project is past 1.0).

Anything intentionally not part of the public API must be **explicitly internal**
(language-specific mechanism: `// @internal`, separate `internal/` subpath, package
private, etc.). If users can import it, they will, and you'll be on the hook for it.

### CHANGELOG is for users, not for you

Write from the user's perspective.

- ❌ `refactor: extract helper function` — users don't care.
- ✅ `feat: add async loading option to loadWorkbook` — users want to know.
- ✅ `fix: loadWorkbook crashed on files with empty sheet names (#123)` — symptom-based.

A pure-internal-refactor PR doesn't need a changeset; if it does need one, mark it
`chore` so it doesn't show up as user-facing.

### Issues and PRs are a conversation, not a queue

- When asked a question, **point at existing docs** before writing code.
- Evaluate feature requests against the "one way to do one thing" rule (is it already
  reachable?).
- For bug reports, **write a failing test first**. Don't fix what you can't reproduce.

See `.claude/skills/issue-triage/SKILL.md` for the full workflow.

## The AI-slop era

**LLM-authored issues, PRs, and review comments are now common.** They tend to be
formally well-structured but substantively thin: not reproducible, already addressed,
out of scope, or copy-pasted from documentation.

This repository takes a hard line:

- **Templates are mandatory.** Issues and PRs that don't follow the templates are
  auto-closed by `.github/workflows/template-compliance.yml`.
- **Repeated low-effort AI submissions can lead to a ban.** This is stated in the
  templates so contributors (and the LLMs they use) know upfront.
- **Using AI is fine. Posting AI output without reading it is not.** Whatever a model
  produces, **a human is responsible** for whether it's worth a maintainer's time.

When Claude or any other LLM works in this repository, it must **re-read its own
output and ask: is this thin, generic, or templated?** before posting anything visible
to the public.

## Skills

The `.claude/skills/` directory contains workflow-specific guides. Use them when the
situation matches.

| Skill                | When to use                                                            |
| -------------------- | ---------------------------------------------------------------------- |
| `pr-workflow`        | Creating a PR                                                          |
| `full-code-review`   | Reviewing a branch from a maintainer's perspective before opening a PR |
| `review-response`    | Responding to GitHub review comments                                   |
| `run-check-and-test` | Running quality checks and tests before commit / PR                    |
| `issue-triage`       | Classifying a GitHub issue and routing it to the right workflow        |

When you add a new skill, append it to this table.

## Operations bundled with the template

- **`SECURITY.md`** — private vulnerability reporting via GitHub Security
  Advisories. Update the support-versions block once the project has a real
  branch policy.
- **`.github/CODEOWNERS`** — review routing. Add per-subsystem entries above
  the wildcard as the project grows.
- **`renovate.json`** — Renovate Bot config. Enable Renovate on the repo
  ([dashboard](https://developer.mend.io/)) to get grouped dependency
  bumps, weekly lock-file maintenance, and security alert labelling.
  Native-module projects should add a `matchPackageNames` rule that
  prevents auto-merge for build-toolchain deps.
