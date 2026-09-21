---
'@office-kit/pptx-dev': minor
---

`init` now writes the project guide as `AGENTS.md` as well as `CLAUDE.md`, so coding agents other than Claude Code pick up the TSX workflow.

Claude Code loads `CLAUDE.md` automatically; Codex and most other agents load `AGENTS.md`. A generated project only had the first, so those agents started with no instructions and reached for their own slide-building tools — producing a deck outside the project instead of editing `deck.tsx`. Both files carry the same text, and existing projects can get the same result by copying `CLAUDE.md` to `AGENTS.md`.
