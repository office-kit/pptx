---
'pptx-kit': minor
---

feat: chart trendline `<c:forward>` / `<c:backward>` extensions.
`ChartTrendline.forward` and `backward` carry the N-period
extrapolation past the last / before the first data point. The
playground renderer projects the linear fit further along the x-axis
by `N * step` so projected-future trendlines render the way
PowerPoint shows them. Moving-average / log / poly trendlines keep
their data-range output since extrapolation isn't meaningful for
them.
