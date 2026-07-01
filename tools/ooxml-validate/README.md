# ooxml-validate

A tiny .NET console app that runs Microsoft's **OpenXmlValidator** (the engine
behind the [Open XML SDK Productivity Tool](https://github.com/dotnet/Open-XML-SDK))
over generated `.pptx` files.

## Why this exists

The test suite already validates every emitted XML part against the ECMA-376
XSDs with `xmllint`. That catches **structural** schema violations, but XSDs
can't express the **semantic** OOXML rules — attribute co-constraints,
relationship targeting, value ranges, part-level invariants — that PowerPoint
enforces when it decides a file is corrupt. `OpenXmlValidator` is a second,
independent oracle that checks those rules.

It is intentionally a separate toolchain (.NET, not the JS hot path), so it
stays out of the published bundle and the per-test loop. It runs in CI over the
`pnpm samples` decks.

## Usage

```sh
# from the repo root, after `pnpm build && pnpm samples`
dotnet run --project tools/ooxml-validate -c Release -- samples/out
```

Pass a directory (scanned recursively for `*.pptx`) or explicit file paths.
Exits `0` when every file is clean, `1` when any validation error is found
(each printed with its Id, description, owning part, and XPath), `2` on bad
invocation.

## Pinning

- **.NET SDK**: `8.0.404` (`global.json`, `rollForward: latestFeature`).
- **DocumentFormat.OpenXml**: `3.1.0` (`ooxml-validate.csproj`).

Both are pinned so the validator's rule set is reproducible across machines and
CI. The validator targets `FileFormatVersions.Microsoft365` — the newest band,
so post-2007 features aren't reported as unsupported.
