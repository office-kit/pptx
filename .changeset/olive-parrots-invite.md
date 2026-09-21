---
'@office-kit/pptx-dsl': major
---

fix: `Text` with `bullets` crashed on `@office-kit/pptx` 0.20.0 with `api.setShapeBullets is not a function`

0.20.0 renamed `setShapeBullets` to `setShapeBulletStyle`. The DSL was updated in
the same change but never republished, because its `>=0.17.0` peer range kept the
new core in range and the release tooling only bumps peer dependents that fall out
of range. npm and pnpm therefore installed the two together and every deck using
`bullets` failed at render time.

The peer range is now `^0.20.0`, so `@office-kit/pptx` 0.17–0.19 are no longer
accepted. Upgrade the core package alongside this one.
