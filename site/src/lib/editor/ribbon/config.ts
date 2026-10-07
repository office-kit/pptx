// Ribbon layout — a PowerPoint-style tab/group/command arrangement over the
// capability manifest. Each command id here must exist in the manifest (guarded
// at load below), but the ribbon is deliberately NOT the coverage surface: any
// capability the ribbon does not list is still reachable through the properties
// panel (auto-generated from the manifest) and the Ctrl+K palette. The ribbon's
// job is ergonomics for the common path, not exhaustiveness.

import { inches } from '@office-kit/pptx';
import { capabilityById } from '../manifest/index.ts';
import type { EditorController } from '../core/controller.svelte.ts';

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
  readonly contextual?: 'shape' | 'cell' | 'image' | 'table' | 'chart' | 'media' | 'master';
  readonly groups: readonly RibbonGroup[];
}

export const RIBBON: readonly RibbonTab[] = [
  {
    id: 'slideMaster',
    title: 'Slide Master',
    contextual: 'master',
    // Mac PowerPoint's Slide Master tab. Edits act on the layout behind the
    // current slide, so every slide sharing it follows.
    groups: [
      {
        title: 'Edit Master',
        items: [
          {
            id: 'insertSlideMaster',
            icon: 'new-slide',
            label: 'Insert Slide Master',
            unavailable: 'Adding masters and layouts is not supported by the library yet.',
          },
          {
            id: 'insertLayout',
            icon: 'slide-content',
            label: 'Insert Layout',
            unavailable: 'Adding masters and layouts is not supported by the library yet.',
          },
          {
            id: 'deleteLayout',
            icon: 'trash',
            label: 'Delete',
            unavailable: 'Deleting layouts is not supported by the library yet.',
          },
          { id: 'setSlideLayoutName', icon: 'rename', label: 'Rename' },
        ],
      },
      {
        title: 'Master Layout',
        items: [{ id: 'setSlideLayoutPlaceholderBounds', icon: 'align', label: 'Master Layout' }],
      },
      {
        title: 'Edit Theme',
        items: [
          { id: 'setPresentationTheme', icon: 'theme', label: 'Colors' },
          { id: 'setPresentationFonts', icon: 'font', label: 'Fonts' },
        ],
      },
      {
        title: 'Background',
        items: [
          { id: 'setSlideMasterBackgroundStyle', icon: 'background', label: 'Background Styles' },
          { id: 'setSlideLayoutBackground', icon: 'background', label: 'Format Background' },
          { id: 'clearSlideLayoutBackground', icon: 'trash', label: 'Reset Background' },
        ],
      },
      { title: 'Size', items: [{ id: 'setSlideSize', icon: 'resize', label: 'Slide Size' }] },
      {
        title: 'Close',
        items: [
          {
            id: 'closeMasterView',
            icon: 'close-master',
            label: 'Close Master',
            run: (editor) => (editor.masterView = false),
            enabled: () => true,
          },
        ],
      },
    ],
  },
  {
    id: 'home',
    title: 'Home',
    // Laid out by HomeRibbon.svelte, which mirrors Mac PowerPoint's clusters.
    groups: [],
  },
  {
    id: 'insert',
    title: 'Insert',
    // Laid out by InsertRibbon.svelte, with Mac PowerPoint's ▾ menus.
    groups: [],
  },
  { id: 'draw', title: 'Draw', groups: [] },
  {
    id: 'design',
    title: 'Design',
    groups: [],
  },
  { id: 'transitions', title: 'Transitions', groups: [] },
  { id: 'animations', title: 'Animations', groups: [] },
  { id: 'slideShow', title: 'Slide Show', groups: [] },
  { id: 'record', title: 'Record', groups: [] },
  { id: 'review', title: 'Review', groups: [] },
  { id: 'view', title: 'View', groups: [] },
  {
    // PowerPoint's Chart Design tab. The chart dialog edits the data, type and
    // elements in one place, so those commands open it.
    id: 'chartDesign',
    title: 'Chart Design',
    contextual: 'chart',
    groups: [
      {
        title: 'Chart Layouts',
        items: [
          { id: 'setChartSpec', icon: 'chart', label: 'Add Chart Element' },
          {
            id: 'setChartSpec',
            icon: 'layout',
            label: 'Quick Layout',
            unavailable: 'Quick layouts are not available in this editor yet.',
          },
        ],
      },
      {
        title: 'Chart Styles',
        items: [
          {
            id: 'setChartSpec',
            icon: 'theme',
            label: 'Change Colors',
            unavailable: 'Change series colors in Edit Data.',
          },
          {
            id: 'setChartSpec',
            icon: 'quick-styles',
            label: 'Chart Styles',
            unavailable: 'Chart styles are not available in this editor yet.',
          },
        ],
      },
      {
        title: 'Data',
        items: [
          {
            id: 'setChartSpec',
            icon: 'swap',
            label: 'Switch Row/Column',
            unavailable: 'Switching rows and columns is not available in this editor yet.',
          },
          { id: 'setChartSpec', icon: 'table', label: 'Select Data' },
          { id: 'setChartSpec', icon: 'table', label: 'Edit Data' },
        ],
      },
      {
        title: 'Type',
        items: [{ id: 'setChartSpec', icon: 'chart', label: 'Change Chart Type' }],
      },
    ],
  },
  { id: 'shape', title: 'Shape Format', contextual: 'shape', groups: [] },
  // Laid out by PictureFormatRibbon, TableDesignRibbon and TableLayoutRibbon.
  { id: 'picture', title: 'Picture Format', contextual: 'image', groups: [] },
  { id: 'tableDesign', title: 'Table Design', contextual: 'table', groups: [] },
  { id: 'table', title: 'Table Layout', contextual: 'table', groups: [] },
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
