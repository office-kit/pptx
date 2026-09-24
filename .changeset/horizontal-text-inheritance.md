---
'@office-kit/pptx': minor
---

Fix horizontal text overrides on placeholders inheriting vertical text from a layout or master. Explicit `horz` values now stop the inheritance cascade. Behavior change: `setShapeTextDirection(shape, 'horz')` writes an explicit override; use `null` to clear the override and restore inheritance.
