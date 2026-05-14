# References

This directory holds material used **only while developing pptx-kit**. None of
it ships in the published npm package (the `files` whitelist in `package.json`
excludes everything outside `dist/`).

## Layout

- `python-pptx/`, `PptxGenJS/`, `pptx-automizer/`, `Open-XML-SDK/`,
  `ShapeCrawler/`, `open-xml-docs/`, `ecma-376-5th/`, `docxtemplater/`,
  `officecrypto-tool/`, `pptx-renderer/` — git submodules pointing at
  permissively-licensed reference implementations. Shallow clones. See
  `.gitmodules` for the canonical URLs.
- `EXTERNAL.md` — repos we deliberately do **not** track as submodules
  (size or copyleft license). Clone them yourself, outside this repo,
  when you need them.
- `specs/` — workspace for ECMA-376 PDFs and XSDs, downloaded by
  `specs/fetch.sh`. The downloaded files themselves are gitignored.

## Why submodules?

Reading the reference implementations is unavoidable for an OOXML library.
Pinning each at a specific commit means our copy doesn't shift under our feet
when we cite "see python-pptx file X line Y."

## Init

After cloning this repo:

```sh
git submodule update --init --recursive --depth 1
```

This populates all 10 submodule directories with shallow clones (~150 MB
total). To skip the references entirely (rare — you'll lose access to
schemas and reference code while developing):

```sh
git config submodule.recurse false
```
