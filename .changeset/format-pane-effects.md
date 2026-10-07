---
'@office-kit/pptx-dev': minor
---

The editor's Format Shape pane now follows Mac PowerPoint more closely:

- **Shape Options / Text Options** switch at the top. Text Options has PowerPoint's Text Fill & Outline, Text Effects and Textbox categories and applies to the text of the selected shapes.
- **Effects** shows PowerPoint's Shadow, Reflection, Glow, Soft Edges, 3-D Format and 3-D Rotation sections with their preset galleries, color buttons, sliders and boxes, instead of the editor's list of effect commands. Text Effects offers the same sections for the text (Soft Edges and Keep text flat are shown disabled, with the reason, because the library cannot write them yet).
- **Every section starts collapsed**, as in PowerPoint, and stays as you left it for the session. Size and Position... still opens the Size and Position sections.
- The Line section gains **Sketched style** (shown disabled: the library does not write sketched lines yet), and the Begin/End Arrow type and size buttons are PowerPoint's 39 pt gallery buttons. Arrow sizes are named Arrow L Size 1–9 / Arrow R Size 1–9 as in PowerPoint.
- The flip checkboxes left the Position section (PowerPoint has none there); Arrange ▸ Rotate keeps Flip Vertical and Flip Horizontal.
