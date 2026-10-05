// Ribbon layout — a PowerPoint-style tab/group/command arrangement over the
// capability manifest. Each command id here must exist in the manifest (guarded
// at load below), but the ribbon is deliberately NOT the coverage surface: any
// capability the ribbon does not list is still reachable through the properties
// panel (auto-generated from the manifest) and the Ctrl+K palette. The ribbon's
// job is ergonomics for the common path, not exhaustiveness.

import { inches } from '@office-kit/pptx';
import { capabilityById } from '../manifest/index.ts';
import type { EditorController } from '../core/controller.svelte.ts';
import { insertMedia, insertScreenshot, insertWordArt } from '../core/insert-objects.ts';
import { t } from '../i18n/i18n.svelte.ts';

// Default drop placement for inserted objects — like PowerPoint dropping a
// default-sized shape you then move/resize. EMU via the public unit helpers.
const IN = (n: number) => inches(n) as unknown as number;
const DROP = { x: IN(2), y: IN(1.5), w: IN(4), h: IN(2) };
export const PRESET = {
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
  /** Runs instead of the capability when the native command does more than one call. */
  readonly run?: (editor: EditorController, button: HTMLElement) => void;
  /** Whether a `run` command applies now (a capability uses `canRun`). */
  readonly enabled?: (editor: EditorController) => boolean;
  /**
   * Why a native command is shown but cannot run here (no recognizer, no
   * library, no web API); the button stays disabled with this as its tip.
   */
  readonly unavailable?: string;
}

export interface RibbonGroup {
  readonly title: string;
  readonly items: readonly RibbonItem[];
}

export interface RibbonTab {
  readonly id: string;
  readonly title: string;
  /** When set, the tab only shows for this selection kind (contextual tab). */
  readonly contextual?: 'shape' | 'cell' | 'image' | 'table' | 'media';
  readonly groups: readonly RibbonGroup[];
}

export const RIBBON: readonly RibbonTab[] = [
  {
    id: 'home',
    title: 'Home',
    // Laid out by HomeRibbon.svelte, which mirrors Mac PowerPoint's clusters.
    groups: [],
  },
  {
    id: 'insert',
    title: 'Insert',
    // Mac PowerPoint's Insert tab, in its order and with its names.
    groups: [
      {
        title: 'Slides',
        items: [
          {
            id: 'addSlide',
            icon: 'new-slide',
            label: 'New Slide',
            run: (editor) => editor.addNewSlide(),
          },
        ],
      },
      { title: 'Tables', items: [{ id: 'addSlideTable', icon: 'table', label: 'Table' }] },
      {
        title: 'Images',
        items: [
          { id: 'addSlideImage', icon: 'picture', label: 'Pictures' },
          {
            id: 'insertScreenshot',
            icon: 'screenshot',
            label: 'Screenshot',
            run: (editor) => void insertScreenshot(editor, t('Screenshot')),
            enabled: (editor) =>
              !!editor.doc.currentSlide &&
              typeof navigator !== 'undefined' &&
              !!navigator.mediaDevices?.getDisplayMedia,
          },
        ],
      },
      {
        title: 'Camera',
        items: [
          {
            id: 'cameo',
            icon: 'cameo',
            label: 'Cameo',
            unavailable: 'Recording is not available in the browser.',
          },
        ],
      },
      {
        title: 'Illustrations',
        items: [
          { id: 'addSlideShape', icon: 'shapes', label: 'Shapes', preset: PRESET.shape },
          {
            id: 'icons',
            icon: 'icons',
            label: 'Icons',
            unavailable: 'The Office icon library is not available here.',
          },
          {
            id: '3dModels',
            icon: 'cube',
            label: '3D Models',
            unavailable: '3D models are not supported by the library yet.',
          },
          {
            id: 'smartArt',
            icon: 'smartart',
            label: 'SmartArt',
            unavailable: 'SmartArt is not supported by the library yet.',
          },
          { id: 'addSlideChart', icon: 'chart', label: 'Chart' },
        ],
      },
      {
        title: 'Links',
        items: [
          {
            id: 'zoom',
            icon: 'zoom-slide',
            label: 'Zoom',
            unavailable: 'Slide zoom is not supported by the library yet.',
          },
          { id: 'setShapeHyperlink', icon: 'link', label: 'Link' },
          { id: 'setShapeClickAction', icon: 'action', label: 'Action' },
        ],
      },
      { title: 'Comments', items: [{ id: 'addSlideComment', icon: 'comment', label: 'Comment' }] },
      {
        title: 'Text',
        items: [
          { id: 'addSlideTextBox', icon: 'textbox', label: 'Text Box', preset: PRESET.textBox },
          {
            id: 'headerFooter',
            icon: 'header-footer',
            label: 'Header & Footer',
            run: (editor) => (editor.activeDialog = 'headerFooter'),
            enabled: (editor) => !!editor.doc.currentSlide,
          },
          {
            id: 'insertWordArt',
            icon: 'wordart',
            label: 'WordArt',
            run: (editor) => insertWordArt(editor, t('WordArt'), t('Your text here')),
            enabled: (editor) => !!editor.doc.currentSlide,
          },
          // Both insert a field PowerPoint keeps up to date into the selected
          // box; Date & Time asks for the format first, as the native dialog does.
          { id: 'setShapeTextField', icon: 'calendar', label: 'Date & Time' },
          {
            id: 'setShapeTextField',
            icon: 'slide-number',
            label: 'Slide Number',
            preset: { type: 'slidenum' },
          },
          {
            id: 'object',
            icon: 'object',
            label: 'Object',
            unavailable: 'Embedded OLE objects are not supported by the library yet.',
          },
        ],
      },
      {
        title: 'Symbols',
        items: [
          {
            id: 'equation',
            icon: 'equation',
            label: 'Equation',
            unavailable: 'Equations are not supported by the library yet.',
          },
          {
            id: 'insertSymbol',
            icon: 'symbol',
            label: 'Symbol',
            // Like PowerPoint, Symbol needs a text cursor to insert at.
            run: (editor, button) => editor.openSymbolPicker(button),
            enabled: (editor) => !!editor.inlineTextFormat?.insertText,
          },
        ],
      },
      {
        title: 'Media',
        items: [
          {
            id: 'insertVideo',
            icon: 'video',
            label: 'Video',
            run: (editor) => void insertMedia(editor, 'video', t('Video')),
            enabled: (editor) => !!editor.doc.currentSlide,
          },
          {
            id: 'insertAudio',
            icon: 'audio',
            label: 'Audio',
            run: (editor) => void insertMedia(editor, 'audio', t('Audio')),
            enabled: (editor) => !!editor.doc.currentSlide,
          },
        ],
      },
    ],
  },
  { id: 'draw', title: 'Draw', groups: [] },
  {
    id: 'design',
    title: 'Design',
    groups: [
      // Mac PowerPoint's Variants group (Colors, Fonts, Background Styles) and
      // Slide Size; there is no Office theme gallery to pick whole themes from.
      {
        title: 'Variants',
        items: [
          { id: 'setPresentationTheme', icon: 'theme', label: 'Colors' },
          { id: 'setPresentationFonts', icon: 'font', label: 'Fonts' },
        ],
      },
      {
        title: 'Background',
        items: [],
      },
      {
        title: 'Customize',
        items: [{ id: 'setSlideSize', icon: 'resize', label: 'Slide Size' }],
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
  { id: 'transitions', title: 'Transitions', groups: [] },
  { id: 'animations', title: 'Animations', groups: [] },
  { id: 'slideShow', title: 'Slide Show', groups: [] },
  { id: 'record', title: 'Record', groups: [] },
  { id: 'review', title: 'Review', groups: [] },
  { id: 'view', title: 'View', groups: [] },
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
  { id: 'playback', title: 'Playback', contextual: 'media', groups: [] },
];

// Guard: every capability-backed ribbon command id must be a real capability.
// Items with `run` or `unavailable` name a native command, not a capability.
for (const tab of RIBBON) {
  for (const group of tab.groups) {
    for (const item of group.items) {
      if (item.run || item.unavailable) continue;
      if (!capabilityById.has(item.id)) {
        throw new Error(
          `Ribbon references unknown capability "${item.id}" (tab ${tab.id} / ${group.title}).`,
        );
      }
    }
  }
}
