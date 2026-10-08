---
'@office-kit/pptx-dev': minor
---

The editor's window now matches the reference desktop app's (Mac) measurements:

- **Zoom percentages** mean what they mean in the reference desktop app (Mac): at 100% a slide point is one screen point (a widescreen slide is 960 px wide; it was 1280 px), so Fit to Window now reads about 120% in a full-size window instead of about 76%. Fit leaves the reference desktop app's 22 pt margin around the slide.
- **Home ribbon**: a 72 pt command row with the reference desktop app's button sizes and group spacing; Add-ins and Designer are separate groups; the second Font row is in the reference desktop app's order (… Character Spacing, Change Case, then Text Highlight Color and Font Color). The extra Font dialog (A…) button is removed — use Cmd+T or Character Spacing ▸ More Spacing....
- **Layout**: the thumbnail pane is 249 pt wide, the notes pane opens one line tall, and the status bar uses the reference desktop app's sizes.
- **Format Shape pane**: a compact title, larger category tabs, chevron section headers, and one-line label / control rows with the reference desktop app's 26 pt controls and 112 pt pop-ups. The Line section now starts with No line / Solid line / Gradient line, replacing the No outline button, and hides its settings for No line.
- **Right-click menus** on objects and slide thumbnails list the reference desktop app's items in its order, with its separators and 24 pt rows. Objects gain Lock/Unlock, Reorder Overlapping Objects and Action Settings...; Link... and Edit Alt Text... are now Hyperlink... and View Alt Text...; commands the editor cannot perform yet are shown disabled with the reason. Thumbnails gain Select All, Zoom... and Slide Show, and no longer list Layout or Reset Slide (they remain on the Home tab).
