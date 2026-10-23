---
'@office-kit/pptx-preview': patch
---

The browser text path (`textLayout: 'foreignObject'`, the default) now sets `text-autospace: no-autospace` on text boxes and table cells. Presentation files have no setting for extra space between Japanese / Chinese and Latin text, and presentation apps add none, but CSS `text-autospace` may default to adding it in browsers. With that default, a browser preview drew lines like 「9 月に 15,640 件」 wider than the `textLayout: 'svg'` path and `auditTextLayout` measured them, and could wrap them differently.
