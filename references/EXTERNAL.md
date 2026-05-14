# External references — clone outside the repo, never vendored

This directory's `EXTERNAL.md` lists reference projects that we want to read
during development but that we deliberately **do not** track as submodules.
Reasons fall into two buckets:

- **Size.** Apache POI and LibreOffice are tens of GB. Submodule init time
  alone would be a contributor onboarding tax.
- **License.** GPL / LGPL / MPL / AGPL projects cannot be safely *copied* into
  this MIT codebase. We're allowed to *read* them, but vendoring invites
  accidental verbatim copying.

If you want any of these locally, clone them next to this repository — `git
clone --depth 1 ...` is enough for reference reading.

| Repository | License | Suggested local path | Why we care |
|---|---|---|---|
| [`apache/poi`](https://github.com/apache/poi) | Apache-2.0 | `~/git/refs/poi` | XSLF is the most battle-tested non-Microsoft PPTX read/write implementation. Schema interpretation we can cross-check. |
| [`plutext/docx4j`](https://github.com/plutext/docx4j) | Apache-2.0 | `~/git/refs/docx4j` | Includes `pptx4j`. JAXB-faithful schema mirror — useful when ECMA wording is ambiguous. |
| [`LibreOffice/core`](https://github.com/LibreOffice/core) | MPL-2.0 / LGPL-3.0 / GPL-3.0 | `~/git/refs/lo-core` | The `oox/` and `sd/` modules handle PowerPoint OOXML import/export. **Copyleft — read only. Do not copy code.** |
| [`ONLYOFFICE/sdkjs`](https://github.com/ONLYOFFICE/sdkjs) | **AGPL-3.0** | `~/git/refs/sdkjs` | Full browser-side PPTX editor. **AGPL — read only. Do not copy code.** Patterns and approach are useful; literal code is not. |

## Quick clone

```sh
# Anywhere outside this repo:
mkdir -p ~/git/refs
cd ~/git/refs
git clone --depth 1 https://github.com/apache/poi.git
git clone --depth 1 https://github.com/plutext/docx4j.git
# Sparse-checkout for LibreOffice — only the oox and sd modules:
git clone --depth 1 --filter=blob:none --no-checkout https://github.com/LibreOffice/core.git lo-core
cd lo-core
git sparse-checkout init --cone
git sparse-checkout set oox sd
git checkout
cd ..
git clone --depth 1 https://github.com/ONLYOFFICE/sdkjs.git
```

## License discipline

When you spot something useful in one of the above:

- **Apache-2.0** (POI, docx4j): you may study the design freely. If you lift
  more than a few lines, add an attribution comment naming the source file
  and commit hash. Apache-2.0 patent grant is permissive enough that this is
  safe.
- **MPL / LGPL / GPL** (LibreOffice): study only. Do not copy code, including
  XML constant tables, into this repo. Write your own implementation.
- **AGPL** (sdkjs): study only. Same rules — and remember AGPL's network-use
  trigger if you ever combine this code with a server.

When in doubt, re-implement from the ECMA-376 spec (Part 1, §19 for
PresentationML) rather than the reference implementation.
