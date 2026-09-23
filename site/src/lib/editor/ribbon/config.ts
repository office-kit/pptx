// Ribbon layout — a PowerPoint-style tab/group/command arrangement over the
// capability manifest. Each command id here must exist in the manifest (guarded
// at load below), but the ribbon is deliberately NOT the coverage surface: any
// capability the ribbon does not list is still reachable through the properties
// panel (auto-generated from the manifest) and the Ctrl+K palette. The ribbon's
// job is ergonomics for the common path, not exhaustiveness.

import { inches } from '@office-kit/pptx';
import { capabilityById } from '../manifest/index.ts';

// Default drop placement for inserted objects — like PowerPoint dropping a
// default-sized shape you then move/resize. EMU via the public unit helpers.
const IN = (n: number) => inches(n) as unknown as number;
const DROP = { x: IN(2), y: IN(1.5), w: IN(4), h: IN(2) };
const PRESET = {
  shape: { opts: { preset: 'rect', x: DROP.x, y: DROP.y, w: DROP.w, h: DROP.h } },
  textBox: { opts: { x: DROP.x, y: DROP.y, w: DROP.w, h: IN(1), text: 'Text' } },
  line: { opts: { from: { x: IN(2), y: IN(3) }, to: { x: IN(7), y: IN(3) } } },
} as const;

export interface RibbonItem {
  /** Capability id to run (via runOrPrompt). */
  readonly id: string;
  /** Optional preset args applied before prompting for the rest. */
  readonly preset?: Record<string, unknown>;
  /** Override label (else the manifest label). */
  readonly label?: string;
  /** Short visible label; the full label stays available to assistive technology. */
  readonly compactLabel?: string;
  readonly icon?: string;
}

export interface RibbonGroup {
  readonly title: string;
  readonly items: readonly RibbonItem[];
}

export interface RibbonTab {
  readonly id: string;
  readonly title: string;
  /** When set, the tab only shows for this selection kind (contextual tab). */
  readonly contextual?: 'shape' | 'cell' | 'image' | 'table';
  readonly groups: readonly RibbonGroup[];
}

export const RIBBON: readonly RibbonTab[] = [
  {
    id: 'home',
    title: 'Home',
    groups: [
      {
        title: 'Slides',
        items: [
          { id: 'addBlankSlide', icon: 'slide-blank' },
          { id: 'addSlide', icon: 'slide-content' },
          { id: 'addTitleSlide', icon: 'slide-title' },
          { id: 'addContentSlide', icon: 'slide-content' },
          { id: 'duplicateSlide', icon: 'duplicate' },
          { id: 'removeSlide', icon: 'trash' },
        ],
      },
      {
        title: 'Layout',
        items: [
          { id: 'resetSlideLayout', icon: 'slide-content' },
          {
            id: 'addMissingSlidePlaceholders',
            icon: 'slide-content',
            compactLabel: 'Restore placeholders',
          },
          {
            id: 'resetSlidePlaceholderTextFormatting',
            icon: 'text-format',
            compactLabel: 'Reset text styles',
          },
          { id: 'resetSlidePlaceholderGeometry', icon: 'align', compactLabel: 'Reset positions' },
        ],
      },
      {
        title: 'Font',
        items: [
          { id: 'setShapeTextFormat', icon: 'text-format', label: 'Text format' },
          { id: 'setShapeRunFormat', icon: 'text-run', label: 'Run format' },
        ],
      },
      {
        title: 'Paragraph',
        items: [
          // The manifest labels these "Set paragraph …", which the ribbon
          // truncates to four indistinguishable buttons. Use PowerPoint's own
          // short names instead.
          { id: 'setParagraphAlignment', icon: 'align', label: 'Align' },
          { id: 'setShapeBulletStyle', icon: 'bullets', label: 'Bullets' },
          { id: 'setParagraphLevel', icon: 'indent', label: 'Indent' },
          { id: 'setParagraphLineSpacing', icon: 'line-spacing', label: 'Line spacing' },
          { id: 'setParagraphSpacing', icon: 'space', label: 'Spacing' },
        ],
      },
      {
        title: 'Drawing',
        items: [
          { id: 'addSlideShape', icon: 'shape', preset: PRESET.shape },
          { id: 'addSlideTextBox', icon: 'textbox', preset: PRESET.textBox },
          { id: 'setShapeFill', icon: 'fill' },
          { id: 'setShapeStroke', icon: 'outline' },
          { id: 'setShapeShadow', icon: 'shadow' },
        ],
      },
      {
        title: 'Arrange',
        items: [
          { id: 'bringShapeToFront', icon: 'front' },
          { id: 'sendShapeToBack', icon: 'back' },
          { id: 'groupShapes', icon: 'group' },
          { id: 'setShapeAlignment', icon: 'align' },
        ],
      },
      {
        title: 'Editing',
        items: [{ id: 'replaceTextInPresentation', icon: 'replace', label: 'Replace' }],
      },
    ],
  },
  {
    id: 'insert',
    title: 'Insert',
    groups: [
      {
        title: 'Tables',
        items: [{ id: 'addSlideTable', icon: 'table' }],
      },
      {
        title: 'Illustrations',
        items: [
          { id: 'addSlideShape', icon: 'shape', preset: PRESET.shape },
          { id: 'addSlideImage', icon: 'image' },
          { id: 'addSlideChart', icon: 'chart' },
          { id: 'addSlideLine', icon: 'line', preset: PRESET.line },
        ],
      },
      {
        title: 'Text',
        items: [
          { id: 'addSlideTextBox', icon: 'textbox', preset: PRESET.textBox },
          // Turns the selected box into a field PowerPoint keeps up to date.
          // The deck-wide slide-number switch lives in the slide panel, since
          // it is one setting for the whole file rather than a shape operation.
          { id: 'setShapeTextField', icon: 'text-format', label: 'Insert field' },
          { id: 'setShapeHyperlink', icon: 'link' },
        ],
      },
      {
        title: 'Comments',
        items: [{ id: 'addSlideComment', icon: 'comment' }],
      },
    ],
  },
  {
    id: 'design',
    title: 'Design',
    groups: [
      {
        title: 'Slide setup',
        items: [{ id: 'setSlideSize', icon: 'resize' }],
      },
      {
        title: 'Background',
        items: [
          { id: 'setSlideBackground', icon: 'background' },
          { id: 'setSlideBackgroundImage', icon: 'image' },
          { id: 'clearSlideBackground', icon: 'trash' },
        ],
      },
      {
        title: 'Theme',
        items: [
          { id: 'setPresentationTheme', icon: 'theme' },
          { id: 'setPresentationFonts', icon: 'font' },
        ],
      },
      {
        // Acts on the layout behind the current slide, so every slide sharing
        // it follows — Google Slides' theme builder, without a separate view.
        title: 'Layout',
        items: [
          { id: 'setSlideLayoutName', icon: 'slide-content', compactLabel: 'Rename layout' },
          { id: 'setSlideLayoutBackground', icon: 'background', compactLabel: 'Layout background' },
          { id: 'clearSlideLayoutBackground', icon: 'trash', compactLabel: 'Reset background' },
          {
            id: 'setSlideLayoutPlaceholderBounds',
            icon: 'align',
            compactLabel: 'Move placeholder',
          },
        ],
      },
    ],
  },
  {
    id: 'transitions',
    title: 'Transitions',
    groups: [
      {
        title: 'Transition',
        items: [
          { id: 'setSlideTransition', icon: 'transition' },
          { id: 'clearSlideTransition', icon: 'trash' },
        ],
      },
    ],
  },
  {
    id: 'animations',
    title: 'Animations',
    groups: [
      {
        title: 'Animation',
        items: [
          { id: 'setShapeAnimation', icon: 'animation' },
          { id: 'clearSlideAnimations', icon: 'trash' },
        ],
      },
    ],
  },
  {
    id: 'shape',
    title: 'Shape Format',
    contextual: 'shape',
    groups: [
      {
        title: 'Fill',
        items: [
          { id: 'setShapeFill', icon: 'fill' },
          { id: 'setShapeGradientFill', icon: 'gradient' },
          { id: 'setShapePatternFill', icon: 'pattern' },
          { id: 'setShapeImageFill', icon: 'image' },
          { id: 'setShapeNoFill', icon: 'no-fill' },
        ],
      },
      {
        title: 'Outline',
        items: [
          { id: 'setShapeStroke', icon: 'outline' },
          { id: 'setShapeStrokeDash', icon: 'dash' },
          { id: 'setShapeStrokeArrow', icon: 'arrow' },
          { id: 'setShapeNoStroke', icon: 'no-fill' },
        ],
      },
      {
        title: 'Effects',
        items: [
          { id: 'setShapeShadow', icon: 'shadow' },
          { id: 'setShapeGlow', icon: 'glow' },
          { id: 'clearShapeEffects', icon: 'trash' },
        ],
      },
      {
        title: 'Size & rotate',
        items: [
          { id: 'setShapeBounds', icon: 'resize' },
          { id: 'setShapeRotation', icon: 'rotate' },
          { id: 'setShapeFlip', icon: 'flip' },
        ],
      },
      {
        title: 'Arrange',
        items: [
          { id: 'bringShapeToFront', icon: 'front' },
          { id: 'bringShapeForward', icon: 'forward' },
          { id: 'sendShapeBackward', icon: 'backward' },
          { id: 'sendShapeToBack', icon: 'back' },
          { id: 'groupShapes', icon: 'group' },
          { id: 'ungroupShapes', icon: 'ungroup' },
        ],
      },
    ],
  },
  {
    id: 'table',
    title: 'Table',
    contextual: 'cell',
    groups: [
      {
        title: 'Rows & columns',
        items: [
          { id: 'insertTableRow', icon: 'cells-row' },
          { id: 'insertTableColumn', icon: 'cells-col' },
          { id: 'removeTableRow', icon: 'cells-row' },
          { id: 'removeTableColumn', icon: 'cells-col' },
          { id: 'mergeTableCells', icon: 'merge' },
        ],
      },
      {
        title: 'Cell',
        items: [
          { id: 'setTableCellFill', icon: 'fill' },
          { id: 'setTableCellBorders', icon: 'border' },
          { id: 'setTableCellText', icon: 'text-format' },
          { id: 'setTableCellAlignment', icon: 'align' },
        ],
      },
      {
        title: 'Table style',
        items: [
          { id: 'setTableStyleId', icon: 'theme' },
          { id: 'setTableColumnWidth', icon: 'cells-col' },
          { id: 'setTableRowHeight', icon: 'cells-row' },
        ],
      },
    ],
  },
];

// Guard: every ribbon command id must be a real capability.
for (const tab of RIBBON) {
  for (const group of tab.groups) {
    for (const item of group.items) {
      if (!capabilityById.has(item.id)) {
        throw new Error(
          `Ribbon references unknown capability "${item.id}" (tab ${tab.id} / ${group.title}).`,
        );
      }
    }
  }
}
