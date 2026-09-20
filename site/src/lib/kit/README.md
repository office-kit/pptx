# Office Kit site shell

Everything in this folder is shared by the three Office Kit sites
(`pptx`, `xlsx`, `docx`). It has no imports from the rest of this site apart
from `$lib/components/Search.svelte`, so it can move to a shared package
without untangling anything.

| File                    | What it is                                                                                                                   |
| ----------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `products.ts`           | The family registry: ids, package names, site and repo URLs. The only place that knows where each product lives.             |
| `KitHeader.svelte`      | Sticky header with the product switcher, the product's own nav, search, theme toggle, and the phone menu.                    |
| `KitFooter.svelte`      | Footer with the product's links, the family's links, and the AI-agent entry points.                                          |
| `FamilyGrid.svelte`     | The three libraries side by side, each in its own accent.                                                                    |
| `KitMark.svelte`        | The stacked-files mark, with the current product on top.                                                                     |
| `ThemeToggle.svelte`    | Light / dark switch. Stores the choice under `office-kit-theme`, shared across the three sites because they share an origin. |
| `InstallCommand.svelte` | Copyable `npm i` chip.                                                                                                       |

The design tokens live in `src/app.css`. A product sets `data-product` on
`<html>` (see `src/app.html`) and `CURRENT_PRODUCT` in `products.ts`; the
accent colour and everything derived from it follow from those two values.

## Using it from another product

1. Copy this folder and `src/app.css`.
2. Set `CURRENT_PRODUCT` in `products.ts` and `data-product` in `app.html`.
3. Pass the product's own nav links to `KitHeader` and `KitFooter` from
   `+layout.svelte`.
