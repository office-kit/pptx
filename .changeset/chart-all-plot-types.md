---
'@office-kit/pptx': minor
'@office-kit/pptx-preview': patch
---

feat: author every chart type — scatter, bubble, radar, stock, surface, the 3-D variants and pie-of-pie

`addSlideChart` / `setChartSpec` now write all sixteen plot types of ECMA-376, and `getShapeChartSpec` reads every one of them back.

- `kind: 'scatter' | 'bubble' | 'radar'` are authorable (they used to throw "read-only"). Scatter and bubble series carry their own `xValues` (and `bubbleSizes`); the embedded workbook lays them out per series so "Edit data" opens onto the right cells.
- New kinds `'stock'` (3 series: high, low, close; or 4 with open first, drawn as candlesticks) and `'surface'` (`surfaceContour` for the top-down contour plot, `surfaceWireframe` for the mesh).
- 3-D is a modifier: `view3D: { rotX, rotY, rightAngleAxes, perspective, depthPercent, heightPercent }` on `bar` / `column` / `line` / `area` / `pie` selects the 3-D element, with `bar3DShape` (cylinder, cone, pyramid, …), `gapDepthPct` and `seriesAxis`. `bubble3D` shades bubbles as spheres.
- `ofPie` turns a pie into a pie-of-pie / bar-of-pie chart.
- Per series: `errorBars` / `xErrorBars` (fixed, percentage, standard deviation, standard error, custom), `fillOpacity`.
- Per chart: `dataTable`, `upDownBars`, `categoryAxisDate` (date axis with time units), `categoryGroupLevels` (multi-level category labels), `categoryAxisScaling` (the x axis of scatter / bubble charts, or a date axis' range), `plotAreaLayout` / `titleLayout` / `legend.layout` (manual placement), `valueAxisLineHidden` / `categoryAxisLineHidden`, `valueAxis.displayUnitsLabel`.
- Data labels gain `showBubbleSize`, `showLegendKey`, `fillColor`, and per-point `text` (a literal label, e.g. naming one scatter point).
- A spec whose fields contradict each other (`view3D` on a scatter chart, error bars on a pie, a date axis with non-numeric categories, …) throws with a message naming the field, instead of writing a chart PowerPoint would repair.

Behavior changes when reading existing decks:

- A stock chart now reads as `kind: 'stock'` (was `'line'`) and a surface chart as `kind: 'surface'` (was `'column'`). `getPresentationChartKindCounts` has the two new keys.
- 3-D charts still read as their flat kind, now with `view3D` set, so `setChartSpec(chart, getShapeChartSpec(shape))` no longer flattens a 3-D chart.
- fix: for scatter / bubble charts `valueAxis` (and the other `valueAxis*` fields) described the **x** axis. They now describe the y axis; the x axis reads into `categoryAxis*` and `categoryAxisScaling`.
- fix: charts with more than 25 series wrote workbook references past column `Z` as invalid cell addresses.

`@office-kit/pptx-preview` draws stock charts as lines and surface charts as columns (as before), and honors `categoryAxisScaling` on scatter / bubble charts.
