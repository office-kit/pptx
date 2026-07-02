---
"pptx-kit": patch
"pptx-kit-preview": patch
---

Relax the `engines.node` floor from `>=24.16.0` to `>=22.18.0` on both `pptx-kit` and `pptx-kit-preview` so the maintained LTS lines — Node 22 and Node 24 — are supported, and restore Node 22 to the CI test matrix. The published runtime bundles are unchanged; the previous floor reflected the dev toolchain's pin and needlessly blocked `pnpm install` (under `engine-strict`) on still-supported LTS releases such as Node 22.x and earlier Node 24 LTS patches (e.g. 24.13.x).
