---
'@office-kit/pptx': minor
'@office-kit/pptx-preview': patch
---

`secondaryValueAxis` now takes its scaling fields directly, the same shape as `valueAxis`: `secondaryValueAxis: { min: 0, max: 100, numberFormat: '0"%"', title: 'Rate' }`. Before, they had to be nested under `scaling` (`secondaryValueAxis: { scaling: { min: 0, max: 100 } }`), so code written like the primary axis failed to type-check with "Object literal may only specify known properties".

The nested `scaling` form still works on write but is deprecated. Combining it with the flat fields throws, because the two would disagree.

**Read shape change:** `getSlideCharts` / `readChartSpec` now return the secondary axis's `min`, `max`, `majorUnit`, `minorUnit`, `numberFormat`, `logBase`, `displayUnits` and `displayUnitsLabel` directly on `secondaryValueAxis`, no longer under `secondaryValueAxis.scaling`. Code that reads `spec.secondaryValueAxis.scaling.max` should read `spec.secondaryValueAxis.max`. A spec read from a file and passed back to `setChartSpec` round-trips unchanged.
