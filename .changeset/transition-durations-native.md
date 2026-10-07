---
'@office-kit/pptx': patch
'@office-kit/pptx-dev': patch
---

Transitions are saved the way Mac PowerPoint saves them:

- **`setSlideTransition` with `durationMs`.** `spd` is now the fastest speed at least as long as the duration (≤ 0.5 s fast, ≤ 0.75 s medium, otherwise slow). `fast` is left out because it is the schema default. A duration equal to its speed's own (500, 750 or 1000 ms) is written as that speed alone, with no `p14:dur` and no `mc:AlternateContent` unless the effect needs one. It reads back as `speed` without `durationMs`. Any other duration is written as before.
- **Editor.** Each Transitions gallery tile now writes the element, attributes and duration Mac PowerPoint 16 saves for it, and the ribbon Duration shows PowerPoint's value. Twenty-three durations changed, for example Reveal 3.40, Curtains 6.00, Honeycomb 4.40, Shape 0.80 and Zoom 0.90. Default directions PowerPoint does not write are no longer written, for example on Split, Reveal, Ripple, Shred, Cube and Fly Through, and Wind, Airplane and Origami no longer write `invX`. Ten Japanese gallery names now match PowerPoint, for example 垂れ幕, 破砕, ハチの巣, 細分, 扉 and 窓.
- **Durations of transitions without `spd`.** The ribbon Duration and the slide-show preview now use the schema default (fast, 0.5 s) for a transition without `spd`. They used 0.75 s before.
- **Effect Options.** Choosing an option keeps a duration that is stored only as a speed. Before, the transition fell back to fast.
