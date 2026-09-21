# AGENTS.md

This repository's instructions live in [CLAUDE.md](CLAUDE.md). **Read it before
changing anything here** — it is not Claude-specific. It carries the project
overview, the OOXML rules that decide what correct output is, the public-API
constraints, and the review standard every change is held to.

The split is only about filenames: Claude Code loads `CLAUDE.md` automatically,
Codex and other agents load `AGENTS.md`. Keeping one file as the source avoids
two copies drifting apart, which is why this one is a pointer rather than a
second copy.

A few things worth knowing before you start:

- **Do not author slide decks with your own presentation tooling.** This project
  is that tooling. Decks are written as TSX against `@office-kit/pptx-dsl` and
  exported with `@office-kit/pptx-dev`; see the "TSX presentation authoring"
  section of CLAUDE.md. A generated deck project carries its own guide under
  both names.
- **The package manager is pnpm**, pinned by `packageManager` in
  `package.json`. Use `pnpm`, not `npm` or `yarn`.
- **Quality gates**, in the order CI runs them:

  ```sh
  pnpm format:check && pnpm lint && pnpm typecheck && pnpm test && pnpm build
  ```

  The DSL package type-checks against the built output of the core and preview
  packages, so `pnpm --filter @office-kit/pptx-dsl typecheck` is only meaningful
  after those have been built. Running it against a stale `dist/` reports fewer
  errors than CI will.

- **User-visible changes need a changeset** (`pnpm changeset`). See the
  "CHANGELOG is for users" section of CLAUDE.md for what belongs in one.
