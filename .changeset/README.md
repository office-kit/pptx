# Changesets

This directory holds [changesets](https://github.com/changesets/changesets) — small
markdown files describing user-visible changes that should appear in `CHANGELOG.md`
on the next release.

## When to add a changeset

Add one in the **same PR** as any change that:

- Adds, removes, or renames a public export.
- Changes observable behavior of an existing export.
- Fixes a bug that users could have hit.
- Bumps a peer dependency requirement.

Pure-internal refactors, build / CI / test changes, and doc-only tweaks do **not**
need a changeset.

## How

```sh
pnpm changeset
```

Pick `patch`, `minor`, or `major` per SemVer, write a one-line user-facing
description (symptom-based for fixes; capability-based for features), and commit
the generated file along with your code change.

When this branch lands on `main`, the release workflow will open a "Version
Packages" PR that consumes the changesets and bumps `package.json` + `CHANGELOG.md`.
Merging that PR publishes to npm.
