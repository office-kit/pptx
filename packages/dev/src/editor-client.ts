import { alignedPositions, distributedPositions, slideExtent } from './shape-arrangement.ts';
import {
  transitionEffects,
  transitionOptions,
  transitionOptionSelected,
  type EffectChoice,
} from './transition-gallery.ts';
import { openShowProperties } from './show-properties-dialog.ts';
import { openCustomShows } from './custom-show-dialog.ts';
import { openActionDialog } from './action-dialog.ts';
import { openLinkDialog } from './link-dialog.ts';
import { previewPictureMenu } from './picture-menu-preview.ts';
import { renderPictureCropOverlay } from './picture-crop-overlay.ts';
import { dragPictureCrop } from './picture-crop.ts';
import { textCaseEdits } from './text-case.ts';
import { openParagraphDialog } from './paragraph-dialog.ts';
import { openFontDialog } from './font-dialog.ts';
import { tableCellShape, expandTableSelection } from './table-cells.ts';
import { textFontFamily } from './text-font.ts';
import { findText, type TextMatch, type TextSearchOptions } from './text-search.ts';
import { confirmSlideSize, openPageSetup, slideSizePresets } from './page-setup.ts';
import { openTablePicker } from './table-insert.ts';
import { snapMove, smartGuideTargets } from './smart-guides.ts';
import {
  shapePoint,
  connectorEndpoints,
  connectorFrame,
  type ConnectorFrame,
} from './connector-geometry.ts';
import { shapeCategories, lineTools, lineTool } from './shape-gallery.ts';
import { getPresetShapePath } from '@office-kit/pptx-preview';
import type { EditorModel, EditorShape, EditCommand, EditorTransitionSound } from './editor.ts';
import {
  createRichTextField,
  fitRichText,
  centerRichText,
  measureRichTextBox,
  resizedTextBounds,
  replaceParagraphRange,
  type RichTextField,
} from './rich-text.ts';
import { overlayStyles } from './editor-ui.ts';
import { createShapeProperties } from './shape-properties.ts';
type MenuItem = {
  label: string;
  disabled?: boolean;
  checked?: boolean;
  icon?: string;
  picture?: { src: string; opacity: number; brightness?: number; contrast?: number };
  shortcut?: string;
  heading?: boolean;
  thumbnail?: {
    width: number;
    height: number;
    boxes: Array<{ x: number; y: number; w: number; h: number }>;
  };
  action?: () => void;
  preview?: () => () => void;
  children?: (MenuItem | null)[];
} | null;
interface OfficeState {
  revision: number;
  index: number;
  presenting: boolean;
  animationPreview: boolean;
  building: boolean;
  error: string | null;
  editor?: EditorModel;
  history?: { undo: number; redo: number };
}
declare global {
  interface Window {
    office: {
      getState(): OfficeState;
      selectSlide(index: number, focus?: boolean): void;
      refresh(): Promise<void>;
      startCustomShow(id: number): void;
      previewTransition(): void;
      previewAnimations(timingId?: string | string[], selectedOnly?: boolean): void;
      stopAnimationPreview(): void;
      menu(items: MenuItem[], x: number, y: number, origin?: Element | null): void;
    };
  }
}
const office = window.office;
const byId = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;
const host = byId('slide'),
  shadow = host.shadowRoot!;
const selected = new Set<number>();
let cropping: number | null = null;
let cropAspectRatio: number | undefined;
let cellSelection: {
  id: number;
  slide: number;
  revision: number;
  range: NonNullable<EditCommand['tableRange']>;
} | null = null;
let pendingTableStyle: {
  slide: string | undefined;
  id: number;
  key: string;
  value: boolean;
} | null = null;
let currentSlide = -1,
  currentRevision = -1,
  layer: HTMLDivElement;
let busy = false,
  tool: string | null = null;
let textEdit: {
  field: RichTextField;
  cell?: { row: number; column: number };
  undoText(redo: boolean): boolean;
  undoCount: number;
  redoCount: number;
  value: string;
  edits: Array<{ start: number; end: number; text: string; format?: EditorShape['format'] }>;
  runs: EditorShape['runs'];
  paragraphs: EditorShape['paragraphs'];
  autoFitParams: EditorShape['autoFitParams'];
  bounds: EditorShape['bounds'];
  insertion: { position: number; format: EditorShape['format'] } | null;
  shape: EditorShape;
  revision: number;
  slide: number;
} | null = null;
let clipboardPending = false;
let clipboard: {
  slide: number;
  ids: number[];
  snapshot: string;
  key: string;
  cut: boolean;
} | null = null;
const model = () => office.getState().editor;
const slideModel = () => model()?.slides[office.getState().index];
function flattenShapes(shapes: EditorShape[]): EditorShape[] {
  return shapes.flatMap((shape) => [shape, ...flattenShapes(shape.children ?? [])]);
}
function visibleShapes(shapes = slideModel()?.shapes ?? []): EditorShape[] {
  return shapes.flatMap((shape) =>
    shape.hidden ? [] : [shape, ...visibleShapes(shape.children ?? [])],
  );
}
const allShapes = () => flattenShapes(slideModel()?.shapes ?? []);
const selection = () => allShapes().filter((s) => selected.has(s.id));
const expandedGroups = new Set<number>();
function localVector(shape: EditorShape, x: number, y: number) {
  const m = shape.parentTransform;
  if (!m) return { x, y };
  const [a, b, c, d] = m,
    det = a * d - b * c;
  return { x: (d * x - c * y) / det, y: (-b * x + a * y) / det };
}
function appendShapeOverlay(shape: EditorShape, element: HTMLElement | SVGElement) {
  if (!shape.parentTransform) {
    layer.append(element);
    return;
  }
  const wrapper = document.createElement('div');
  const [a, b, c, d, e, f] = shape.parentTransform;
  wrapper.style.cssText = `position:absolute;inset:0;pointer-events:none;transform-origin:0 0;transform:matrix(${a},${b},${c},${d},${e * ratio()},${f * ratio()})`;
  wrapper.append(element);
  layer.append(wrapper);
}
const error = (message: string) => {
  byId('edit-message').textContent = message;
};
const color = (value: string | null | undefined, fallback: string) =>
  /^#?[0-9a-f]{6}$/i.test(value ?? '') ? '#' + value!.replace(/^#/, '') : fallback;
const ratio = () => host.clientWidth / (model()?.width ?? 1);
const guideSettings = { smart: true, drawing: false, grid: false, snap: false };
let drawingPreference: boolean | undefined;
let defaultGridSnap = false;
let defaultGridSpacing = { x: 72008, y: 72008 };
const gridSpacing = () => model()?.gridSpacing ?? defaultGridSpacing;
let saveGridDefault = false;
try {
  const saved = JSON.parse(localStorage.getItem('office-guide-settings') ?? '{}');
  const defaults = JSON.parse(localStorage.getItem('office-grid-defaults') ?? 'null');
  if (
    defaults &&
    [defaults.x, defaults.y].every((n) => Number.isSafeInteger(n) && n > 0 && n <= 36000000)
  )
    defaultGridSpacing = { x: defaults.x, y: defaults.y };
  if (typeof defaults?.snap === 'boolean') defaultGridSnap = defaults.snap;
  if (typeof saved?.drawing === 'boolean') drawingPreference = saved.drawing;
  for (const key of ['smart', 'grid'] as const)
    if (typeof saved?.[key] === 'boolean') guideSettings[key] = saved[key];
} catch {
  // Storage can be unavailable in private/embedded browsing contexts.
}
type DrawingGuide = { id: number; axis: 'x' | 'y'; offset: number; color: string };
let drawingGuides: DrawingGuide[] = [];
let guideDrag: {
  guide: DrawingGuide;
  node: HTMLElement;
  pointer: number;
  offset: number;
  start: number;
  tooltip: HTMLElement;
} | null = null;
function saveDrawingGuides() {
  if (JSON.stringify(drawingGuides) === JSON.stringify(model()?.guides)) return;
  void send({
    type: 'guides',
    slide: office.getState().index,
    guides: drawingGuides.map((guide) => ({ ...guide })),
  });
}
function guidePosition(guide: DrawingGuide) {
  return ((guide.axis === 'x' ? model()?.width : model()?.height) ?? 0) / 2 + guide.offset;
}
function positionDrawingGuide(node: HTMLElement, guide: DrawingGuide) {
  node.style[guide.axis === 'x' ? 'left' : 'top'] = guidePosition(guide) * ratio() - 3 + 'px';
}
function addDrawingGuide(axis: 'x' | 'y') {
  if (busy) return;
  const same = drawingGuides.filter((guide) => guide.axis === axis);
  drawingGuides.push({
    id: Math.max(0, ...drawingGuides.map((guide) => guide.id)) + 1,
    axis,
    offset: same.length ? same[same.length - 1]!.offset + 360000 : 0,
    color: '#888888',
  });
  setGuideSetting('drawing', true);
  saveDrawingGuides();
  renderDrawingGuides();
}
function finishGuideDrag(cancel: boolean) {
  if (!guideDrag) return;
  const session = guideDrag;
  guideDrag = null;
  session.tooltip.remove();
  if (cancel) session.guide.offset = session.offset;
  else {
    const extent = (session.guide.axis === 'x' ? model()?.width : model()?.height) ?? 0;
    const position = guidePosition(session.guide);
    if (position < 0 || position > extent)
      drawingGuides = drawingGuides.filter((guide) => guide !== session.guide);
    saveDrawingGuides();
  }
  if (host.hasPointerCapture(session.pointer)) host.releasePointerCapture(session.pointer);
  renderDrawingGuides();
}
function renderDrawingGuides() {
  if (guideDrag) finishGuideDrag(true);
  layer?.querySelectorAll('.drawing-guide,.drawing-grid').forEach((node) => node.remove());
  if (layer && guideSettings.grid && !office.getState().presenting) {
    const grid = document.createElement('div');
    grid.className = 'drawing-grid';
    grid.setAttribute('aria-hidden', 'true');
    const spacing = gridSpacing();
    const x = Math.max(2, spacing.x * ratio()),
      y = Math.max(2, spacing.y * ratio());
    grid.style.cssText = `position:absolute;inset:0;pointer-events:none;background-image:radial-gradient(circle at 0.5px 0.5px,#888 0.65px,transparent 0.85px);background-size:${x}px ${y}px`;
    layer.prepend(grid);
  }
  byId<HTMLInputElement>('show-guides').checked = guideSettings.drawing;
  if (!layer || !guideSettings.drawing || office.getState().presenting) return;
  for (const guide of drawingGuides) {
    const node = document.createElement('div');
    node.className = 'drawing-guide';
    node.dataset.axis = guide.axis;
    node.dataset.guideId = String(guide.id);
    node.setAttribute('aria-label', guide.axis === 'x' ? 'Vertical Guide' : 'Horizontal Guide');
    node.style.cssText = `position:absolute;pointer-events:auto;z-index:1;${guide.axis === 'x' ? 'top:0;bottom:0;width:7px;cursor:ew-resize' : 'left:0;right:0;height:7px;cursor:ns-resize'}`;
    const line = document.createElement('div');
    line.style.cssText = `position:absolute;pointer-events:none;${guide.axis === 'x' ? 'left:3px;top:0;bottom:0;border-left' : 'top:3px;left:0;right:0;border-top'}:1px dashed ${guide.color}`;
    node.append(line);
    positionDrawingGuide(node, guide);
    node.onpointerdown = (event) => {
      if (event.button !== 0 || busy || drag || textEdit || tool) return;
      event.preventDefault();
      event.stopPropagation();
      focusCanvas();
      const tooltip = document.createElement('div');
      tooltip.className = 'guide-distance';
      tooltip.style.cssText =
        'position:absolute;pointer-events:none;z-index:3;padding:3px 5px;background:#ffffdf;color:#222;border:1px solid #888;font:11px Arial;white-space:nowrap';
      layer.append(tooltip);
      guideDrag = {
        guide,
        node,
        pointer: event.pointerId,
        offset: guide.offset,
        start: guide.axis === 'x' ? event.clientX : event.clientY,
        tooltip,
      };
      updateGuideDistance(event);
      host.setPointerCapture(event.pointerId);
    };
    node.oncontextmenu = (event) => {
      event.preventDefault();
      event.stopPropagation();
      if (busy) return;
      office.menu(
        [
          { label: 'Add Vertical Guide', action: () => addDrawingGuide('x') },
          { label: 'Add Horizontal Guide', action: () => addDrawingGuide('y') },
          {
            label: 'Delete',
            action: () => {
              drawingGuides = drawingGuides.filter((item) => item !== guide);
              saveDrawingGuides();
              renderDrawingGuides();
            },
          },
          {
            label: 'Color',
            children: [
              ['Gray', '#888888'],
              ['Red', '#c43c3c'],
              ['Orange', '#e48312'],
              ['Yellow', '#c4a300'],
              ['Green', '#25833a'],
              ['Blue', '#2873c4'],
              ['Purple', '#8946ad'],
            ].map(([label, value]) => ({
              label: label!,
              checked: guide.color === value,
              action: () => {
                guide.color = value!;
                saveDrawingGuides();
                renderDrawingGuides();
              },
            })),
          },
        ],
        event.clientX,
        event.clientY,
        host,
      );
    };
    layer.append(node);
  }
}
function updateGuideDistance(event: PointerEvent) {
  if (!guideDrag) return;
  const rect = host.getBoundingClientRect();
  guideDrag.tooltip.textContent = (Math.abs(guideDrag.guide.offset) / 360000).toFixed(2) + ' cm';
  guideDrag.tooltip.style.left = event.clientX - rect.left + 12 + 'px';
  guideDrag.tooltip.style.top = event.clientY - rect.top + 12 + 'px';
}
host.addEventListener(
  'pointermove',
  (event) => {
    if (!guideDrag || guideDrag.pointer !== event.pointerId) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    const position = guideDrag.guide.axis === 'x' ? event.clientX : event.clientY;
    guideDrag.guide.offset = guideDrag.offset + (position - guideDrag.start) / ratio();
    positionDrawingGuide(guideDrag.node, guideDrag.guide);
    updateGuideDistance(event);
  },
  true,
);
for (const type of ['pointerup', 'pointercancel', 'lostpointercapture'] as const)
  host.addEventListener(
    type,
    (event) => {
      if (!guideDrag || guideDrag.pointer !== event.pointerId) return;
      event.stopImmediatePropagation();
      finishGuideDrag(type !== 'pointerup');
    },
    true,
  );
document.addEventListener(
  'keydown',
  (event) => {
    if (!guideDrag) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    if (event.key === 'Escape') finishGuideDrag(true);
  },
  true,
);
function setGuideSetting(key: keyof typeof guideSettings, value: boolean) {
  if (busy) return;
  guideSettings[key] = value;
  if (key === 'snap') {
    void send({ type: 'grid-snap', slide: office.getState().index, snapToGrid: value });
    return;
  }
  if (key === 'drawing') drawingPreference = value;
  try {
    localStorage.setItem(
      'office-guide-settings',
      JSON.stringify({
        smart: guideSettings.smart,
        grid: guideSettings.grid,
        ...(drawingPreference === undefined ? {} : { drawing: drawingPreference }),
      }),
    );
  } catch {
    // Keep the setting usable for this session even without persistent storage.
  }
  layer?.querySelectorAll('.smart-guide').forEach((node) => node.remove());
  renderDrawingGuides();
}
// Native guide visibility follows the application across already-open documents.
window.addEventListener('storage', (event) => {
  if (event.key !== 'office-guide-settings' && event.key !== null) return;
  try {
    const saved = JSON.parse(event.newValue ?? '{}');
    guideSettings.smart = typeof saved?.smart === 'boolean' ? saved.smart : true;
    guideSettings.grid = typeof saved?.grid === 'boolean' ? saved.grid : false;
    drawingPreference = typeof saved?.drawing === 'boolean' ? saved.drawing : undefined;
    guideSettings.drawing = drawingPreference ?? model()?.guidesVisible ?? false;
    layer?.querySelectorAll('.smart-guide').forEach((node) => node.remove());
    renderDrawingGuides();
  } catch {
    // Ignore malformed preferences written by another tab.
  }
});
function openGridOptions() {
  if (busy) return;
  saveGridDefault = false;
  byId<HTMLInputElement>('grid-snap').checked = guideSettings.snap;
  byId<HTMLInputElement>('grid-visible').checked = guideSettings.grid;
  byId<HTMLInputElement>('grid-guides').checked = guideSettings.drawing;
  byId<HTMLInputElement>('grid-smart').checked = guideSettings.smart;
  const spacing = gridSpacing().x;
  const select = byId<HTMLSelectElement>('grid-spacing');
  const preset = [...select.options].find(
    (option) => Math.abs(Number(option.value) - spacing) < 20,
  );
  select.value = preset?.value ?? 'custom';
  byId<HTMLInputElement>('grid-custom').value = String(spacing / 360000);
  byId('grid-custom-row').hidden = select.value !== 'custom';
  byId<HTMLDialogElement>('grid-dialog').showModal();
}
byId('grid-default').onclick = () => {
  saveGridDefault = true;
};
byId('grid-spacing').onchange = () => {
  byId('grid-custom-row').hidden = byId<HTMLSelectElement>('grid-spacing').value !== 'custom';
};
byId('grid-dialog').querySelector('form')!.onsubmit = async (event) => {
  if ((event.submitter as HTMLButtonElement)?.value !== 'apply') return;
  event.preventDefault();
  if (busy) return;
  const choice = byId<HTMLSelectElement>('grid-spacing').value;
  const spacing = Math.round(
    choice === 'custom'
      ? Number(byId<HTMLInputElement>('grid-custom').value) * 360000
      : Number(choice),
  );
  const drawing = byId<HTMLInputElement>('grid-guides').checked;
  const snap = byId<HTMLInputElement>('grid-snap').checked;
  const current = gridSpacing();
  if (
    ((choice === 'custom' ? current.x !== spacing : Math.abs(current.x - spacing) >= 20) ||
      (choice === 'custom' ? current.y !== spacing : Math.abs(current.y - spacing) >= 20) ||
      snap !== guideSettings.snap) &&
    !(await send({
      type: 'grid-spacing',
      slide: office.getState().index,
      gridSpacing: { x: spacing, y: spacing },
      snapToGrid: snap,
    }))
  )
    return;
  setGuideSetting('grid', byId<HTMLInputElement>('grid-visible').checked);
  setGuideSetting('drawing', drawing);
  setGuideSetting('smart', byId<HTMLInputElement>('grid-smart').checked);
  if (saveGridDefault) {
    defaultGridSpacing = { ...gridSpacing() };
    defaultGridSnap = snap;
    try {
      localStorage.setItem(
        'office-grid-defaults',
        JSON.stringify({ ...defaultGridSpacing, snap: defaultGridSnap }),
      );
    } catch {
      /* Keep defaults available for this session when storage is unavailable. */
    }
  }
  byId<HTMLDialogElement>('grid-dialog').close();
};
function gridMenu(): MenuItem[] {
  return [
    {
      label: 'Smart Guides',
      checked: guideSettings.smart,
      action: () => setGuideSetting('smart', !guideSettings.smart),
    },
    {
      label: 'Guides',
      checked: guideSettings.drawing,
      action: () => setGuideSetting('drawing', !guideSettings.drawing),
    },
    {
      label: 'Gridlines',
      checked: guideSettings.grid,
      action: () => setGuideSetting('grid', !guideSettings.grid),
    },
    null,
    {
      label: 'Snap to Grid',
      checked: guideSettings.snap,
      action: () => setGuideSetting('snap', !guideSettings.snap),
    },
    null,
    { label: 'Grid Options...', action: openGridOptions },
  ];
}
function guideMenu(): MenuItem[] {
  return [
    { label: 'Add Vertical Guide', action: () => addDrawingGuide('x') },
    { label: 'Add Horizontal Guide', action: () => addDrawingGuide('y') },
    null,
    {
      label: 'Guides',
      checked: guideSettings.drawing,
      action: () => setGuideSetting('drawing', !guideSettings.drawing),
    },
    {
      label: 'Smart Guides',
      checked: guideSettings.smart,
      action: () => setGuideSetting('smart', !guideSettings.smart),
    },
    null,
    { label: 'Grid Options...', action: openGridOptions },
  ];
}
byId<HTMLInputElement>('show-guides').onchange = (event) =>
  setGuideSetting('drawing', (event.target as HTMLInputElement).checked);
byId('view-grid-guides').onclick = (event) => {
  const button = event.currentTarget as HTMLElement;
  const rect = button.getBoundingClientRect();
  office.menu(guideMenu(), rect.left, rect.bottom, button);
};
const shapeProperties = createShapeProperties({
  selection,
  disabled: () => busy || office.getState().building,
  update: (extra) => perform('update', extra),
  focusCanvas,
});
const elementFor = (id: number) =>
  shadow.querySelector<SVGGElement>(`[data-pptx-shape-id="${id}"]`);
function setBox(element: HTMLElement, box: NonNullable<EditorShape['bounds']>) {
  const scale = ratio();
  Object.assign(element.style, {
    left: box.x * scale + 'px',
    top: box.y * scale + 'px',
    width: box.w * scale + 'px',
    height: box.h * scale + 'px',
  });
}
function focusCanvas() {
  host.tabIndex = 0;
  host.focus({ preventScroll: true });
}
async function send(
  command?: EditCommand,
  action?: 'undo' | 'redo',
  revision = office.getState().revision,
) {
  if (busy) return false;
  office.stopAnimationPreview();
  busy = true;
  error('');
  updateControls();
  try {
    await dragPreviewTask;
    if (command && ['guides', 'guides-visible'].includes(command.type))
      command = { ...command, snapToGrid: guideSettings.snap };
    if (command) command = await measureCommandText(command, revision);
    const response = await fetch('/edit', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ revision, command, action }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error ?? 'Could not save the edit.');
    await new Promise((done) => setTimeout(done, 100));
    // Wait until replay has completed; a second action must use the new revision.
    for (let attempt = 0; attempt < 100; attempt++) {
      await office.refresh();
      const state = office.getState();
      if (state.revision > revision && !state.building) return true;
      if (state.error) throw new Error(state.error);
      await new Promise((done) => setTimeout(done, 40));
    }
    throw new Error(
      'Saving is taking longer than expected. Wait for the presentation to finish updating.',
    );
  } catch (cause) {
    error(cause instanceof Error ? cause.message : String(cause));
    return false;
  } finally {
    busy = false;
    render();
  }
}
async function measureCommandText(command: EditCommand, revision: number): Promise<EditCommand> {
  if (
    !['update', 'text-replace-all'].includes(command.type) ||
    command.textFits ||
    command.changes?.autoFitParams
  )
    return command;
  const changes = command.changes ?? {};
  // Text entry already measures its own live surface. Pure moves/rotations do not affect fit.
  if (changes.textEdits || changes.text !== undefined) return command;
  const source = (model()?.slides ?? []).flatMap((slide, index) =>
    flattenShapes(slide.shapes).map((shape) => ({ shape, slide: index })),
  );
  const replaced =
    command.type === 'text-replace-all' && command.replacement
      ? new Set(
          findText(
            model()?.slides ?? [],
            command.replacement.query,
            command.replacement.options,
          ).map((match) => match.slide + ':' + match.shape),
        )
      : null;
  const candidates = source.filter(({ shape, slide }) => {
    if (!shape.textable || !shape.bounds) return false;
    if (replaced) return replaced.has(slide + ':' + shape.id) && shape.autoFit !== 'none';
    if (slide !== command.slide || !command.ids?.includes(shape.id)) return false;
    const mode = changes.autoFit ?? shape.autoFit;
    if (mode === 'none') return false;
    const next = changes.bounds ?? command.positions?.find((item) => item.id === shape.id)?.bounds;
    return (
      (next && (next.w !== shape.bounds.w || next.h !== shape.bounds.h)) ||
      [
        'autoFit',
        'format',
        'align',
        'bullets',
        'lineSpacing',
        'indent',
        'paragraph',
        'margins',
        'wrap',
        'direction',
        'columns',
        'textColumns',
        'changeCase',
      ].some((key) => Object.hasOwn(changes, key))
    );
  });
  if (!candidates.length) return command;
  const response = await fetch('/edit', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ revision, command, preview: true }),
  });
  const result = await response.json();
  if (!response.ok) throw new Error(result.error ?? 'Could not measure the text change.');
  const preview = result.editor as EditorModel;
  const textFits: NonNullable<EditCommand['textFits']> = [];
  for (const { shape: original, slide } of candidates) {
    const shape = preview.slides[slide]?.shapes.find((item) => item.id === original.id);
    if (!shape?.bounds) continue;
    const field = createRichTextField();
    const scale = ratio();
    field.className = 'direct-text';
    field.value = shape.text;
    setBox(field, shape.bounds);
    styleTextField(field, { ...shape, autoFitParams: null }, scale);
    field.style.visibility = 'hidden';
    field.contentEditable = 'false';
    field.setAttribute('aria-hidden', 'true');
    appendShapeOverlay(shape, field);
    try {
      if (shape.autoFit === 'normal') {
        textFits.push({
          id: shape.id,
          ...(replaced ? { slide } : {}),
          autoFitParams: fitRichText(
            field,
            shape.runs,
            scale,
            shape.paragraphs,
            shape.format.size ?? 18,
            shape.autoFitParams?.lnSpcReduction ?? 0,
          ),
        });
      } else if (shape.autoFit === 'shape') {
        field.renderRuns(shape.runs, scale, shape.paragraphs);
        textFits.push({
          id: shape.id,
          ...(replaced ? { slide } : {}),
          bounds: measuredTextBounds(field, shape, scale),
        });
      }
    } finally {
      field.dispose();
      field.remove();
    }
  }
  return { ...command, textFits };
}
function command(type: EditCommand['type'], extra: Partial<EditCommand> = {}) {
  return {
    type,
    slide: office.getState().index,
    ids: [...selected],
    ...(type === 'order' ? { stableOrder: true } : {}),
    ...(type === 'duplicate' ? { duplicateInParent: true } : {}),
    ...extra,
  };
}
async function clearSelectedContent() {
  const cells = cellSelection;
  if (!cells) return perform('delete');
  if (!(await perform('table-clear', { tableRange: cells.range }))) return false;
  if (office.getState().index === cells.slide && selected.size === 1 && selected.has(cells.id)) {
    cellSelection = { ...cells, revision: office.getState().revision };
    render();
    focusCanvas();
  }
  return true;
}
async function perform(type: EditCommand['type'], extra: Partial<EditCommand> = {}) {
  if (type === 'slide-add' && extra.layout === undefined) {
    const layouts = model()?.layouts ?? [];
    const current = layouts.find((layout) => layout.key === slideModel()?.layout);
    const next =
      current?.type === 'title'
        ? (layouts.find((layout) => layout.type === 'obj') ?? current)
        : current;
    if (next) extra = { ...extra, layout: next.key };
  }
  const selectedCells = cellSelection;
  const rangeFormatting =
    selectedCells &&
    ((type === 'update' &&
      extra.changes &&
      Object.keys(extra.changes).every((key) =>
        [
          'format',
          'align',
          'bullets',
          'lineSpacing',
          'anchor',
          'direction',
          'paragraph',
          'indent',
          'changeCase',
        ].includes(key),
      )) ||
      type === 'table-fill' ||
      type === 'table-borders' ||
      type === 'table-cell-size' ||
      type === 'table-distribute' ||
      type === 'table-margins');
  if (rangeFormatting) {
    if (type === 'update') type = 'table-range-format';
    extra = { ...extra, tableRange: selectedCells.range };
  }
  const paragraphFormat =
    extra.changes &&
    ['align', 'bullets', 'lineSpacing', 'indent', 'paragraph'].some((key) =>
      Object.hasOwn(extra.changes!, key),
    );
  const cellLayout =
    textEdit?.cell &&
    (extra.changes?.anchor !== undefined || extra.changes?.direction !== undefined);
  const range =
    ((type === 'update' &&
      (extra.changes?.format || extra.changes?.changeCase || paragraphFormat || cellLayout)) ||
      type === 'table-style' ||
      type === 'table-fill' ||
      type === 'table-borders' ||
      type === 'table-distribute' ||
      type === 'table-cell-size' ||
      type === 'table-margins') &&
    textEdit
      ? { start: textEdit.field.selectionStart, end: textEdit.field.selectionEnd }
      : null;
  if (range && range.start === range.end && textEdit && extra.changes?.format && !paragraphFormat) {
    textEdit.insertion = {
      position: range.start,
      format: { ...textEdit.insertion?.format, ...extra.changes!.format },
    };
    textEdit.field.focus();
    updateControls();
    return true;
  }
  const restoredRange =
    range && textEdit && extra.changes?.changeCase
      ? {
          start: range.start,
          end:
            range.end +
            textCaseEdits(textEdit.field.value, extra.changes.changeCase, range).reduce(
              (delta, edit) => delta + edit.text.length - (edit.end - edit.start),
              0,
            ),
        }
      : range;
  const editingCell = textEdit?.cell;
  if (
    (type === 'table-fill' ||
      type === 'table-borders' ||
      type === 'table-distribute' ||
      type === 'table-cell-size' ||
      type === 'table-margins') &&
    editingCell
  )
    extra = { ...extra, cell: { ...editingCell, edits: [] } };
  const editingId = textEdit?.shape.id;
  if (!(await finishText())) return false;
  if (range) extra = { ...extra, changes: { ...extra.changes, range } };
  if (
    range &&
    editingCell &&
    (extra.changes?.format ||
      extra.changes?.changeCase ||
      extra.changes?.paragraph ||
      extra.changes?.indent !== undefined ||
      extra.changes?.align ||
      extra.changes?.lineSpacing !== undefined ||
      extra.changes?.bullets !== undefined ||
      cellLayout)
  ) {
    type = extra.changes?.changeCase
      ? 'table-cell-case'
      : extra.changes?.indent !== undefined
        ? 'table-cell-indent'
        : extra.changes?.paragraph
          ? 'table-cell-paragraph'
          : cellLayout
            ? 'table-cell-layout'
            : extra.changes?.bullets !== undefined
              ? 'table-cell-bullets'
              : extra.changes?.lineSpacing !== undefined
                ? 'table-cell-line-spacing'
                : extra.changes?.align
                  ? 'table-cell-align'
                  : 'table-cell-format';
    extra = { ...extra, cell: { ...editingCell, edits: [] } };
  }
  const before = allShapes().map((s) => s.id);
  const ungroupedIds = new Set(
    type === 'ungroup' ? selection().flatMap((s) => (s.children ?? []).map((c) => c.id)) : [],
  );
  const editedSlideKey = slideModel()?.key;
  const position = office.getState().index;
  if (!(await send(command(type, extra)))) return false;
  // Navigation remains available while the edit is saved. Its completion must
  // not select unrelated objects or move focus on the newly displayed slide.
  if (slideModel()?.key !== editedSlideKey) return true;
  if (type === 'slide-add' || type === 'slide-duplicate') office.selectSlide(position + 1, true);
  if (['insert', 'duplicate', 'paste', 'group', 'ungroup', 'regroup'].includes(type)) {
    selected.clear();
    const shapes = allShapes();
    const added = new Set(shapes.filter((s) => !before.includes(s.id)).map((s) => s.id));
    for (const shape of shapes) {
      if (ungroupedIds.has(shape.id)) selected.add(shape.id);
      else if (
        added.has(shape.id) &&
        !shapes.some(
          (parent) =>
            added.has(parent.id) &&
            flattenShapes(parent.children ?? []).some((child) => child.id === shape.id),
        )
      )
        selected.add(shape.id);
    }
    render();
    focusCanvas();
    if (type === 'insert' && extra.preset?.startsWith('actionButton')) {
      void action('action-settings', host);
    }
    if (type === 'insert' && extra.preset === 'text') {
      const shape = selection()[0];
      if (shape) startText(shape);
    }
  }
  if (rangeFormatting && selectedCells && selected.size === 1 && selected.has(selectedCells.id)) {
    cellSelection = { ...selectedCells, revision: office.getState().revision };
    render();
    focusCanvas();
  }
  if (range && editingId !== undefined) {
    const shape = allShapes().find((s) => s.id === editingId);
    if (shape) {
      if (editingCell)
        startCell(shape, editingCell.row, editingCell.column, restoredRange ?? range);
      else startText(shape, restoredRange ?? range);
    }
  }
  return true;
}
function activeShape() {
  let shape = selection()[0];
  if (shape?.table && cellSelection?.id === shape.id) {
    const { row, column } = cellSelection.range;
    shape = tableCellShape(shape, row, column) ?? shape;
  }
  if (!shape || !textEdit || textEdit.shape.id !== shape.id) return shape;
  const start = textEdit.field.selectionStart;
  const caret = textEdit.field.selectionStart === textEdit.field.selectionEnd;
  const position = caret && start > 0 ? start - 1 : start;
  const run =
    textEdit.runs.find((r) => r.start <= position && position < r.end) ?? textEdit.runs.at(-1);
  const paragraph = textEdit.paragraphs.find((p) => p.start <= start && start <= p.end);
  return {
    ...textEdit.shape,
    align: paragraph ? paragraph.align : shape.align,
    bullets: paragraph ? paragraph.bullets : shape.bullets,
    format: { ...(run?.format ?? shape.format), ...(caret ? textEdit.insertion?.format : {}) },
  };
}
let pendingTransitionDuration = false;
let selectedAnimation: { slide: string; id: string; ids?: string[]; anchor?: string } | null = null;
function selectedAnimationIds(): string[] {
  if (!selectedAnimation || selectedAnimation.slide !== slideModel()?.key) return [];
  const existing = new Set(
    slideModel()
      ?.animations.flat()
      .map((effect) => effect.timingId),
  );
  return (selectedAnimation.ids ?? [selectedAnimation.id]).filter((id) => existing.has(id));
}
function selectAnimation(
  id: string,
  modifiers: { shiftKey: boolean; metaKey: boolean; ctrlKey: boolean },
) {
  const slide = slideModel();
  if (!slide) return;
  const order = slide.animations
    .flat()
    .map((item) => item.timingId)
    .filter((item): item is string => !!item);
  const current = selectedAnimationIds();
  const anchor = selectedAnimation?.anchor ?? selectedAnimation?.id ?? id;
  let ids = [id];
  if (modifiers.shiftKey && order.includes(anchor)) {
    const from = order.indexOf(anchor),
      to = order.indexOf(id);
    ids = order.slice(Math.min(from, to), Math.max(from, to) + 1);
    if (modifiers.metaKey || modifiers.ctrlKey) ids = [...new Set([...current, ...ids])];
  } else if (modifiers.metaKey || modifiers.ctrlKey) {
    ids = current.includes(id) ? current.filter((item) => item !== id) : [...current, id];
  }
  selectedAnimation = ids.length
    ? {
        slide: slide.key,
        id: ids.includes(id) ? id : ids.at(-1)!,
        ids,
        anchor: modifiers.shiftKey ? anchor : id,
      }
    : null;
  updateAnimationControls();
}
let pendingAnimationTiming = false;
let animationListSignature = '';
function activeAnimation() {
  if (selectedAnimation?.slide !== slideModel()?.key) return undefined;
  return slideModel()
    ?.animations.flat()
    .find((effect) => effect.timingId === selectedAnimation?.id);
}
function canMoveAnimation(direction: 'earlier' | 'later') {
  if (busy || pendingAnimationTiming || !activeAnimation()) return false;
  const effects = slideModel()?.animations.flat() ?? [];
  const ids = new Set(selectedAnimationIds());
  return effects.some((effect, index) => {
    const neighbor = effects[index + (direction === 'earlier' ? -1 : 1)];
    return ids.has(effect.timingId ?? '') && neighbor && !ids.has(neighbor.timingId ?? '');
  });
}
function renderAnimationMarkers() {
  if (!layer) return;
  const focused = shadow.activeElement as HTMLElement | null;
  const focusedId = focused?.classList.contains('animation-marker')
    ? focused.dataset.animationId
    : null;
  const focusedShape = focused?.dataset.animationShape;
  layer.querySelector('.animation-markers')?.remove();
  const slide = slideModel();
  if (
    !slide ||
    office.getState().presenting ||
    tool ||
    cropping !== null ||
    (byId('tab-animations').getAttribute('aria-selected') !== 'true' &&
      byId('animation-pane').hidden)
  )
    return;
  const markers = document.createElement('div');
  markers.className = 'animation-markers';
  const stacks = new Map<string, number>();
  const scale = ratio();
  const shapes = visibleShapes(slide.shapes);
  slide.animations.forEach((group, index) =>
    group.forEach((effect) => {
      if (!effect.timingId) return;
      for (const id of effect.shapeIds) {
        const shape = shapes.find((item) => String(item.id) === id);
        const box = shape && slideExtent(shape);
        if (!shape || !box) continue;
        const offset = stacks.get(id) ?? 0;
        stacks.set(id, offset + 1);
        const marker = document.createElement('button');
        marker.type = 'button';
        marker.className = 'animation-marker';
        marker.dataset.animationId = effect.timingId;
        marker.dataset.animationShape = id;
        marker.textContent = String(
          index + (slide.animations[0]?.[0]?.trigger === 'clickEffect' ? 1 : 0),
        );
        marker.setAttribute(
          'aria-label',
          `Animation ${index + (slide.animations[0]?.[0]?.trigger === 'clickEffect' ? 1 : 0)}: ${shape.name}`,
        );
        marker.setAttribute(
          'aria-pressed',
          String(selectedAnimationIds().includes(effect.timingId!)),
        );
        marker.disabled = busy || pendingAnimationTiming;
        marker.style.left = box.x * scale - 25 + 'px';
        marker.style.top = box.y * scale + offset * 21 + 'px';
        marker.onpointerdown = (event) => event.stopPropagation();
        marker.onclick = async (event) => {
          event.stopPropagation();
          if (
            busy ||
            pendingAnimationTiming ||
            !(await finishText()) ||
            slideModel()?.key !== slide.key
          )
            return;
          selectedAnimation = { slide: slide.key, id: effect.timingId! };
          selected.clear();
          for (const target of effect.shapeIds) {
            const item = shapes.find((candidate) => String(candidate.id) === target);
            if (item) selected.add(item.id);
          }
          render();
          [...layer.querySelectorAll<HTMLButtonElement>('.animation-marker')]
            .find(
              (button) =>
                button.dataset.animationId === effect.timingId &&
                button.dataset.animationShape === id,
            )
            ?.focus({ preventScroll: true });
        };
        marker.onkeydown = (event) => {
          if (event.key === 'Delete' || event.key === 'Backspace') {
            event.preventDefault();
            event.stopPropagation();
            selectedAnimation = { slide: slide.key, id: effect.timingId! };
            void action('animation-remove', marker);
          }
        };
        markers.append(marker);
      }
    }),
  );
  layer.append(markers);
  if (focusedId)
    [...markers.querySelectorAll<HTMLButtonElement>('button')]
      .find(
        (button) =>
          button.dataset.animationId === focusedId &&
          button.dataset.animationShape === focusedShape,
      )
      ?.focus({ preventScroll: true });
}
let animationDrag: { slide: string; ids: string[] } | null = null;
function clearAnimationDrop() {
  byId('animation-list')
    .querySelectorAll<HTMLElement>('[data-drop-position]')
    .forEach((node) => {
      delete node.dataset.dropPosition;
      node.style.boxShadow = '';
    });
}
async function dropAnimations(beforeId: string | null) {
  const drag = animationDrag;
  animationDrag = null;
  clearAnimationDrop();
  if (!drag || busy || pendingAnimationTiming || slideModel()?.key !== drag.slide) return;
  const order = slideModel()!
    .animations.flat()
    .map((effect) => effect.timingId!);
  if (beforeId && drag.ids.includes(beforeId)) return;
  const rest = order.filter((id) => !drag.ids.includes(id));
  const at = beforeId === null ? rest.length : rest.indexOf(beforeId);
  if (at < 0) return;
  const desired = [
    ...rest.slice(0, at),
    ...order.filter((id) => drag.ids.includes(id)),
    ...rest.slice(at),
  ];
  if (order.every((id, index) => id === desired[index])) return;
  if (!(await finishText()) || slideModel()?.key !== drag.slide) return;
  if (await perform('animation-reorder', { animationIds: drag.ids, animationBeforeId: beforeId })) {
    selectedAnimation = { slide: drag.slide, id: drag.ids[0]!, ids: drag.ids };
    updateControls();
    [...byId('animation-list').querySelectorAll<HTMLButtonElement>('button')]
      .find((button) => button.dataset.animationId === drag.ids[0])
      ?.focus();
  }
}

const animationPane = byId('animation-pane');
function animationTailDrop(event: DragEvent) {
  if (!animationDrag || busy || pendingAnimationTiming || slideModel()?.key !== animationDrag.slide)
    return null;
  const list = byId('animation-list');
  const last = list.lastElementChild as HTMLElement | null;
  if (
    !last ||
    (event.target !== list && event.target !== animationPane) ||
    event.clientY < last.getBoundingClientRect().bottom
  )
    return null;
  return last;
}
animationPane.addEventListener('dragover', (event) => {
  const last = animationTailDrop(event);
  if (!last) return;
  event.preventDefault();
  if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';
  clearAnimationDrop();
  last.dataset.dropPosition = 'after';
  last.style.boxShadow = 'inset 0 -2px #c43e1c';
});
animationPane.addEventListener('drop', (event) => {
  if (!animationTailDrop(event)) return;
  event.preventDefault();
  event.stopPropagation();
  void dropAnimations(null);
});
animationPane.addEventListener('dragleave', (event) => {
  if (!(event.relatedTarget instanceof Node) || !animationPane.contains(event.relatedTarget))
    clearAnimationDrop();
});

function animationContextMenu(id: string, point?: { x: number; y: number }) {
  if (busy || pendingAnimationTiming) return;
  if (!selectedAnimationIds().includes(id))
    selectAnimation(id, { shiftKey: false, metaKey: false, ctrlKey: false });
  const anchor = [...byId('animation-list').querySelectorAll<HTMLButtonElement>('button')].find(
    (button) => button.dataset.animationId === id,
  );
  if (!anchor) return;
  anchor.focus();
  const ids = new Set(selectedAnimationIds());
  const effects =
    slideModel()
      ?.animations.flat()
      .filter((effect) => ids.has(effect.timingId ?? '')) ?? [];
  const starts = [
    ['clickEffect', 'Start On Click'],
    ['withEffect', 'Start With Previous'],
    ['afterEffect', 'Start After Previous'],
  ] as const;
  const items: MenuItem[] = starts.map(([start, label]) => ({
    label,
    checked: effects.length > 0 && effects.every((effect) => effect.trigger === start),
    action: () => void changeAnimationStart(start),
  }));
  items.push({ label: 'Remove', action: () => void action('animation-remove', anchor) });
  if (point) office.menu(items, point.x, point.y, anchor);
  else menuAt(items, anchor);
}

function updateAnimationControls() {
  const slide = slideModel();
  const effect = activeAnimation();
  if (!effect) selectedAnimation = null;
  const ids = selectedAnimationIds();
  const effects =
    slide?.animations.flat().filter((item) => ids.includes(item.timingId ?? '')) ?? [];
  const common = (key: 'effect' | 'trigger' | 'durationMs' | 'delayMs') =>
    effects.every((item) => item[key] === effect?.[key]);
  document.querySelectorAll<HTMLButtonElement>('[data-edit="animation-add"]').forEach((button) => {
    button.disabled = busy || pendingAnimationTiming || (!effect && !selected.size);
    button.setAttribute(
      'aria-pressed',
      String(common('effect') && effect?.effect === button.dataset.effect),
    );
  });
  const previewSelected = document.querySelector<HTMLButtonElement>(
    '[data-edit="animation-preview-selected"]',
  );
  if (previewSelected) previewSelected.disabled = busy || pendingAnimationTiming || !effect;
  renderAnimationMarkers();
  for (const direction of ['earlier', 'later'] as const)
    document
      .querySelectorAll<HTMLButtonElement>(`[data-edit="animation-${direction}"]`)
      .forEach((button) => {
        button.disabled = !canMoveAnimation(direction);
      });
  byId<HTMLButtonElement>('animation-remove').disabled = busy || pendingAnimationTiming || !effect;
  const start = byId<HTMLSelectElement>('animation-start');
  start.disabled = busy || pendingAnimationTiming || !effect;
  if (!pendingAnimationTiming) start.value = common('trigger') ? (effect?.trigger ?? '') : '';
  const input = byId<HTMLInputElement>('animation-duration');
  input.disabled =
    busy ||
    pendingAnimationTiming ||
    !effect ||
    effects.some(
      (item) => !['fadeIn', 'fadeOut'].includes(item.effect ?? '') || item.durationMs === null,
    );
  if (document.activeElement !== input && !pendingAnimationTiming)
    input.value =
      !common('durationMs') || effect?.durationMs == null ? '' : String(effect.durationMs / 1000);
  const delay = byId<HTMLInputElement>('animation-delay');
  delay.disabled = busy || pendingAnimationTiming || !effect;
  if (document.activeElement !== delay && !pendingAnimationTiming)
    delay.value = effect && common('delayMs') ? String((effect.delayMs ?? 0) / 1000) : '';
  const list = byId('animation-list');
  const signature = JSON.stringify([
    slide?.key,
    slide?.animations,
    flattenShapes(slide?.shapes ?? []).map((s) => [s.id, s.name]),
    selectedAnimation,
    busy,
  ]);
  if (signature === animationListSignature) return;
  animationListSignature = signature;
  const focusId = (document.activeElement as HTMLElement | null)?.dataset.animationId;
  list.replaceChildren();
  slide?.animations.forEach((group, index) =>
    group.forEach((item) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.setAttribute('role', 'option');
      button.setAttribute(
        'aria-selected',
        String(selectedAnimationIds().includes(item.timingId ?? '')),
      );
      button.disabled = busy || !item.timingId;
      button.dataset.animationId = item.timingId ?? '';
      const names = item.shapeIds.map(
        (id) => flattenShapes(slide.shapes).find((shape) => String(shape.id) === id)?.name ?? id,
      );
      const label = item.effect
        ? { appear: 'Appear', disappear: 'Disappear', fadeIn: 'Fade', fadeOut: 'Fade' }[item.effect]
        : 'Effect';
      button.textContent = `${index + (slide?.animations[0]?.[0]?.trigger === 'clickEffect' ? 1 : 0)}  ${item.presetClass === 'exit' ? '↗' : '★'}  ${names.join(', ')} — ${label ?? 'Effect'}`;
      button.draggable = !!item.timingId && !busy;
      button.ondragstart = (event) => {
        if (!item.timingId || busy || pendingAnimationTiming || !event.dataTransfer) {
          event.preventDefault();
          return;
        }
        const ids = selectedAnimationIds().includes(item.timingId)
          ? selectedAnimationIds()
          : [item.timingId];
        animationDrag = { slide: slide.key, ids };
        selectedAnimation = { slide: slide.key, id: item.timingId, ids };
        list
          .querySelectorAll<HTMLButtonElement>('button')
          .forEach((row) =>
            row.setAttribute('aria-selected', String(ids.includes(row.dataset.animationId!))),
          );
        event.dataTransfer.effectAllowed = 'move';
        event.dataTransfer.setData('application/x-office-animation', item.timingId);
      };
      button.ondragover = (event) => {
        if (!animationDrag || animationDrag.slide !== slide.key || busy) return;
        event.preventDefault();
        if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';
        clearAnimationDrop();
        const box = button.getBoundingClientRect();
        const after = event.clientY >= box.top + box.height / 2;
        button.dataset.dropPosition = after ? 'after' : 'before';
        button.style.boxShadow = after ? 'inset 0 -2px #c43e1c' : 'inset 0 2px #c43e1c';
      };
      button.ondrop = (event) => {
        if (!animationDrag) return;
        event.preventDefault();
        event.stopPropagation();
        const order = slide.animations.flat().map((effect) => effect.timingId!);
        const box = button.getBoundingClientRect();
        const after = event.clientY >= box.top + box.height / 2;
        const before = after ? (order[order.indexOf(item.timingId!) + 1] ?? null) : item.timingId!;
        void dropAnimations(before);
      };
      button.ondragend = () => {
        animationDrag = null;
        clearAnimationDrop();
        updateAnimationControls();
      };
      button.onclick = (event) => {
        if (!item.timingId || busy) return;
        selectAnimation(item.timingId, event);
      };
      button.oncontextmenu = (event) => {
        event.preventDefault();
        event.stopPropagation();
        if (item.timingId)
          animationContextMenu(item.timingId, { x: event.clientX, y: event.clientY });
      };
      button.onkeydown = (event) => {
        if (event.key === 'ContextMenu' || (event.shiftKey && event.key === 'F10')) {
          event.preventDefault();
          event.stopPropagation();
          if (item.timingId) animationContextMenu(item.timingId);
          return;
        }
        if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'a') {
          event.preventDefault();
          event.stopPropagation();
          const ids = slide.animations
            .flat()
            .map((item) => item.timingId)
            .filter((id): id is string => !!id);
          selectedAnimation = { slide: slide.key, id: item.timingId!, ids, anchor: ids[0]! };
          updateAnimationControls();
          return;
        }
        if (event.key === 'Delete' || event.key === 'Backspace') {
          event.preventDefault();
          event.stopPropagation();
          void action('animation-remove', button);
          return;
        }
        const buttons = [...list.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')];
        const position = buttons.indexOf(button);
        const target =
          event.key === 'ArrowDown'
            ? buttons[position + 1]
            : event.key === 'ArrowUp'
              ? buttons[position - 1]
              : event.key === 'Home'
                ? buttons[0]
                : event.key === 'End'
                  ? buttons.at(-1)
                  : undefined;
        if (target) {
          event.preventDefault();
          event.stopPropagation();
          target.focus();
          selectAnimation(target.dataset.animationId!, event);
        }
      };
      list.append(button);
    }),
  );
  if (focusId)
    [...list.querySelectorAll<HTMLButtonElement>('button')]
      .find((button) => button.dataset.animationId === focusId)
      ?.focus();
}
let loadingTransitionSound = false;
let pendingSound: { key: string; value: EditorTransitionSound } | null = null;
let soundFileSlide: string | undefined;
let pendingAdvanceTiming: {
  slide: string | undefined;
  advanceOnClick: boolean;
  advanceAfterMs: number | null;
} | null = null;
function updateControls() {
  updateAnimationControls();
  const sound =
    pendingSound?.key === slideModel()?.key ? pendingSound?.value : slideModel()?.transitionSound;
  const soundSelect = byId<HTMLSelectElement>('transition-sound');
  soundSelect.querySelector('[value="embedded"]')?.remove();
  if (sound?.kind === 'play') soundSelect.add(new Option(sound.name, 'embedded'), 2);
  soundSelect.value = sound?.kind === 'play' ? 'embedded' : (sound?.kind ?? 'none');
  soundSelect.disabled = busy || loadingTransitionSound || !!pendingSound || !slideModel();
  const soundLoop = byId<HTMLInputElement>('transition-sound-loop');
  soundLoop.checked = sound?.kind === 'play' && sound.loop;
  soundLoop.disabled = soundSelect.disabled || sound?.kind !== 'play';
  const durationInput = byId<HTMLInputElement>('transition-duration');
  durationInput.disabled =
    busy ||
    pendingTransitionDuration ||
    !!pendingAdvanceTiming ||
    !slideModel() ||
    !slideModel()?.transition ||
    slideModel()?.transition?.effect === 'none';
  if (document.activeElement !== durationInput && !pendingTransitionDuration) {
    const duration = slideModel()?.transition?.durationMs;
    durationInput.value = duration === undefined ? '' : String(duration / 1000);
  }
  const pending = pendingAdvanceTiming?.slide === slideModel()?.key ? pendingAdvanceTiming : null;
  const timing = pending
    ? {
        advanceOnClick: pending.advanceOnClick,
        advanceAfterMs: pending.advanceAfterMs ?? undefined,
      }
    : slideModel()?.transition;
  const after = timing?.advanceAfterMs;
  byId<HTMLInputElement>('advance-click').checked = timing?.advanceOnClick !== false;
  byId<HTMLInputElement>('advance-after').checked = after !== undefined;
  for (const id of ['advance-click', 'advance-after'])
    byId<HTMLInputElement>(id).disabled =
      busy || pendingTransitionDuration || !!pendingAdvanceTiming || !slideModel();
  const time = byId<HTMLInputElement>('advance-time');
  time.disabled =
    busy ||
    pendingTransitionDuration ||
    !!pendingAdvanceTiming ||
    !slideModel() ||
    after === undefined;
  if (document.activeElement !== time) {
    const hundredths = Math.round((after ?? 0) / 10);
    time.value =
      String(Math.floor(hundredths / 6000)).padStart(2, '0') +
      ':' +
      ((hundredths % 6000) / 100).toFixed(2).padStart(5, '0');
  }

  shapeProperties.render();
  const state = office.getState(),
    shapes = selection(),
    shape = activeShape();
  const hasTable = shapes.length === 1 && !!shape?.table;
  const hasPictures = shapes.length > 0 && shapes.every((s) => s.kind === 'picture');
  const mergeRange = cellSelection?.range;
  for (const id of ['tab-table-design', 'tab-table-layout', 'tab-picture-format']) {
    const visible = id === 'tab-picture-format' ? hasPictures : hasTable;
    const tableTab = document.getElementById(id)!;
    tableTab.hidden = !visible;
    if (!visible && tableTab.getAttribute('aria-selected') === 'true') {
      const hadFocus = tableTab === document.activeElement;
      document.getElementById('tab-home')!.click();
      if (hadFocus) document.getElementById('tab-home')!.focus();
    }
  }
  for (const input of document.querySelectorAll<HTMLInputElement>('[data-table-style]')) {
    const key = input.dataset.tableStyle as keyof NonNullable<EditorShape['table']>['style'];
    input.checked =
      pendingTableStyle?.slide === slideModel()?.key &&
      pendingTableStyle?.id === shape?.id &&
      pendingTableStyle?.key === key
        ? pendingTableStyle.value
        : !!shape?.table?.style[key];
    input.disabled = busy || !!pendingTableStyle || !hasTable;
  }
  for (const axis of ['rows', 'columns'] as const) {
    const input = byId<HTMLInputElement>('table-size-' + axis);
    input.disabled = busy || !hasTable;
    if (document.activeElement === input) continue;
    const table = shape?.table;
    const values = table ? (axis === 'rows' ? table.rowHeights : table.columnWidths) : [];
    const range = cellSelection?.range;
    const cell = textEdit?.cell ?? (range ? { row: range.row, column: range.column } : undefined);
    const span = cell && table?.cells[cell.row]?.[cell.column]?.span;
    const start = cell ? (axis === 'rows' ? cell.row : cell.column) : 0;
    const single =
      !!span && (!range || (range.rows === span.rowSpan && range.columns === span.gridSpan));
    const count = range
      ? axis === 'rows'
        ? range.rows
        : range.columns
      : span
        ? axis === 'rows'
          ? span.rowSpan
          : span.gridSpan
        : values.length;
    const chosen = values.slice(start, start + count);
    const total = values.reduce((sum, value) => sum + value, 0);
    const extent = axis === 'rows' ? shapes[0]?.bounds?.h : shapes[0]?.bounds?.w;
    const mixed = !single && chosen.some((value) => value !== chosen[0]);
    const size = single ? chosen.reduce((sum, value) => sum + value, 0) : (chosen[0] ?? 0);
    input.value =
      !mixed && total && extent
        ? String(Number(((size * extent) / total / 914400).toFixed(3)))
        : '';
    input.placeholder = mixed ? 'Mixed' : '';
  }
  const editable = shapes.length > 0 && shapes.every((s) => s.textable);
  for (const element of document.querySelectorAll<HTMLButtonElement>('[data-edit]')) {
    const action = element.dataset.edit!;
    if (action.startsWith('animation-preview')) {
      element.disabled =
        action === 'animation-preview-stop'
          ? !office.getState().animationPreview
          : busy ||
            pendingAnimationTiming ||
            !slideModel()?.animations?.length ||
            (action === 'animation-preview-selected' && !activeAnimation());
      continue;
    }
    if (action === 'animation-earlier' || action === 'animation-later') {
      element.disabled = !canMoveAnimation(action === 'animation-earlier' ? 'earlier' : 'later');
      continue;
    }
    if (action === 'animation-remove') {
      element.disabled = busy || pendingAnimationTiming || !activeAnimation();
      continue;
    }
    if (action === 'animation-pane') {
      element.disabled = !slideModel();
      element.setAttribute('aria-expanded', String(!byId('animation-pane').hidden));
      continue;
    }
    if (action === 'animation-add') {
      element.disabled = busy || pendingAnimationTiming || (!activeAnimation() && !shapes.length);
      continue;
    }
    if (action.startsWith('transition-')) {
      element.disabled =
        busy ||
        pendingTransitionDuration ||
        !!pendingAdvanceTiming ||
        !slideModel() ||
        (action === 'transition-options' &&
          !transitionOptions(slideModel()?.transition?.effect ?? 'none').length);
      if (action === 'transition-effect')
        element.setAttribute(
          'aria-pressed',
          String(element.dataset.effect === (slideModel()?.transition?.effect ?? 'none')),
        );
      continue;
    }
    if (action === 'link' || action === 'action-settings') {
      element.disabled = busy || shapes.length !== 1 || !!cellSelection;
      continue;
    }
    element.disabled =
      busy ||
      (action === 'undo'
        ? !(state.history?.undo || textEdit?.undoCount)
        : action === 'redo'
          ? !(state.history?.redo || textEdit?.redoCount)
          : false);
    if (
      [
        'bold',
        'italic',
        'underline',
        'align-left',
        'align-center',
        'align-right',
        'font-grow',
        'font-shrink',
        'strike',
        'superscript',
        'subscript',
        'character-spacing',
        'highlight',
        'bullets',
        'numbering',
        'promote',
        'demote',
        'line-spacing',
        'columns',
        'align-justify',
        'align-distribute',
        'text-direction',
        'text-anchor',
        'change-case',
      ].includes(action)
    )
      element.disabled ||= !(
        editable ||
        ((textEdit?.cell || cellSelection) &&
          [
            'bold',
            'italic',
            'underline',
            'font-grow',
            'font-shrink',
            'strike',
            'superscript',
            'subscript',
            'character-spacing',
            'highlight',
            'align-left',
            'align-center',
            'align-right',
            'align-justify',
            'align-distribute',
            'line-spacing',
            'bullets',
            'numbering',
            'promote',
            'demote',
            'change-case',
            'text-anchor',
            'text-direction',
          ].includes(action))
      );
    if (
      action === 'table-shading' ||
      action === 'table-margins' ||
      action === 'table-borders' ||
      action.startsWith('table-insert-') ||
      action === 'table-delete' ||
      action.startsWith('table-distribute-')
    )
      element.disabled ||= !hasTable;
    if (action === 'table-select')
      element.disabled ||= !hasTable || !(textEdit?.cell || cellSelection);
    if (action === 'table-split') {
      const target =
        textEdit?.cell ?? (mergeRange && { row: mergeRange.row, column: mergeRange.column });
      const span = target && shape?.table?.cells[target.row]?.[target.column]?.span;
      element.disabled ||=
        !hasTable ||
        !span ||
        !!(
          mergeRange &&
          (mergeRange.rows !== span.rowSpan || mergeRange.columns !== span.gridSpan)
        );
    }
    if (action === 'table-merge')
      element.disabled ||=
        !hasTable ||
        !mergeRange ||
        mergeRange.rows * mergeRange.columns < 2 ||
        (shape?.table?.cells.flatMap((row, r) =>
          row.filter(
            (cell, c) =>
              r >= mergeRange.row &&
              r < mergeRange.row + mergeRange.rows &&
              c >= mergeRange.column &&
              c < mergeRange.column + mergeRange.columns &&
              !cell.span.hMerge &&
              !cell.span.vMerge,
          ),
        ).length ?? 0) < 2;
    if (['cut', 'copy'].includes(action))
      element.disabled ||=
        clipboardPending ||
        (textEdit ? textEdit.field.selectionStart === textEdit.field.selectionEnd : !shapes.length);
    if (['picture-transparency', 'picture-corrections'].includes(action))
      element.disabled ||= !hasPictures;
    if (action === 'picture-replace') element.disabled ||= !hasPictures;
    if (action === 'picture-crop-menu') element.disabled ||= !hasPictures;
    if (action === 'picture-crop') {
      element.disabled ||= !hasPictures || shapes.length !== 1;
      element.setAttribute('aria-pressed', String(cropping !== null));
    }
    if (action === 'arrange') element.disabled ||= !shapes.length;
    if (action === 'paste')
      element.disabled ||=
        clipboardPending || (textEdit ? !navigator.clipboard?.readText : !clipboard);
    if (['slide-delete', 'slide-duplicate', 'shapes', 'text', 'picture', 'table'].includes(action))
      element.disabled ||= !slideModel();
    if (['strike', 'superscript', 'subscript', 'bullets', 'numbering'].includes(action))
      element.setAttribute(
        'aria-pressed',
        String(
          shapes.length > 0 &&
            (shapes.length === 1 && shape ? [shape] : shapes).every((s) =>
              action === 'strike'
                ? !!s.format.strike && s.format.strike !== 'noStrike'
                : action === 'superscript'
                  ? (s.format.baseline ?? 0) > 0
                  : action === 'subscript'
                    ? (s.format.baseline ?? 0) < 0
                    : s.bullets === (action === 'bullets' ? 'bullet' : 'number'),
            ),
        ),
      );
    if (['bold', 'italic', 'underline'].includes(action))
      element.setAttribute(
        'aria-pressed',
        String(
          action === 'underline'
            ? !!shape?.format.underline && shape.format.underline !== 'none'
            : !!shape?.format[action as 'bold' | 'italic'],
        ),
      );
  }
  for (const id of ['font-name', 'font-size', 'font-color'])
    byId<HTMLInputElement>(id).disabled = busy || !(editable || textEdit?.cell || cellSelection);
  byId<HTMLInputElement>('shape-fill').disabled =
    busy || !shapes.length || shapes.some((s) => s.kind !== 'shape');
  if (document.activeElement?.id !== 'font-name')
    byId<HTMLInputElement>('font-name').value = shape?.format.font ?? 'Aptos';
  if (document.activeElement?.id !== 'font-size')
    byId<HTMLInputElement>('font-size').value = String(shape?.format.size ?? 18);
  byId<HTMLInputElement>('font-color').value = color(shape?.format.color, '#000000');
  byId<HTMLInputElement>('shape-fill').value = color(shape?.fill, '#4472c4');
}
function render() {
  clearPictureMenuPreview?.();
  const state = office.getState();
  if (!busy && !guideDrag) {
    drawingGuides = (state.editor?.guides ?? []).map((guide) => ({ ...guide }));
    guideSettings.drawing = drawingPreference ?? state.editor?.guidesVisible ?? false;
    guideSettings.snap = state.editor?.snapToGrid ?? defaultGridSnap;
  }
  if (currentSlide !== state.index) {
    if (textEdit) void finishText();
    selected.clear();
    cropping = null;
    currentSlide = state.index;
    currentRevision = -1;
    tool = null;
  }
  if (cropping !== null && (selected.size !== 1 || !selected.has(cropping))) cropping = null;
  if (cropping === null) cropAspectRatio = undefined;
  const exists = new Set(allShapes().map((s) => s.id));
  for (const id of selected) if (!exists.has(id)) selected.delete(id);
  if (
    cellSelection &&
    (cellSelection.slide !== state.index ||
      cellSelection.revision !== state.revision ||
      selected.size !== 1 ||
      !selected.has(cellSelection.id))
  )
    cellSelection = null;
  if (textEdit && shadow.contains(textEdit.field)) {
    updateControls();
    return;
  }
  shadow.querySelector('.editor-layer')?.remove();
  shadow.querySelector('#editor-css')?.remove();
  const style = document.createElement('style');
  style.id = 'editor-css';
  style.textContent = overlayStyles;
  shadow.append(style);
  layer = document.createElement('div');
  layer.className = 'editor-layer';
  layer.hidden = state.presenting || state.animationPreview;
  shadow.append(layer);
  renderDrawingGuides();
  if (!state.presenting)
    for (const shape of visibleShapes()) {
      if (!shape.bounds || shape.parentTransform === null) continue;
      if (shape.parentTransform && !selected.has(shape.id)) continue;
      const hit = document.createElement('div');
      hit.className = 'shape-hit' + (selected.has(shape.id) ? ' selected' : '');
      hit.classList.toggle('picture-cropping', cropping === shape.id);
      hit.dataset.shapeId = String(shape.id);
      hit.setAttribute('aria-label', shape.name);
      hit.setAttribute('role', 'graphics-symbol');
      setBox(hit, shape.bounds);
      if (shape.kind === 'connector') {
        hit.classList.add('connector-hit');
        if (shape.bounds.w * ratio() < 8) {
          hit.style.width = '8px';
          hit.style.marginLeft = `${(shape.bounds.w * ratio() - 8) / 2}px`;
        }
        if (shape.bounds.h * ratio() < 8) {
          hit.style.height = '8px';
          hit.style.marginTop = `${(shape.bounds.h * ratio() - 8) / 2}px`;
        }
      }
      hit.style.transform = `rotate(${shape.rotation}deg)`;
      if (shape.table && cellSelection?.id === shape.id) {
        const { row, column, rows, columns } = cellSelection.range;
        const widths = shape.table.columnWidths,
          heights = shape.table.rowHeights;
        const sum = (values: number[]) => values.reduce((a, b) => a + b, 0);
        const overlay = document.createElement('div');
        overlay.className = 'table-cell-selection';
        overlay.setAttribute('aria-label', `${rows} rows, ${columns} columns selected`);
        Object.assign(overlay.style, {
          position: 'absolute',
          pointerEvents: 'none',
          background: '#4d90fe55',
          outline: '1px solid #3875d7',
          left: (sum(widths.slice(0, column)) / sum(widths)) * 100 + '%',
          top: (sum(heights.slice(0, row)) / sum(heights)) * 100 + '%',
          width: (sum(widths.slice(column, column + columns)) / sum(widths)) * 100 + '%',
          height: (sum(heights.slice(row, row + rows)) / sum(heights)) * 100 + '%',
        });
        hit.append(overlay);
      }
      if (shape.textable && shape.placeholder && !shape.text) {
        const prompt = document.createElement('span');
        prompt.className = 'placeholder-prompt';
        const title = ['title', 'ctrTitle'].includes(shape.placeholder);
        prompt.textContent = title
          ? 'Click to add title'
          : shape.placeholder === 'subTitle'
            ? 'Click to add subtitle'
            : 'Click to add text';
        prompt.style.fontSize = (shape.format.size ?? (title ? 32 : 24)) * 12700 * ratio() + 'px';
        prompt.style.fontFamily = textFontFamily(shape.format.font);
        prompt.style.textAlign =
          shape.align === 'ctr' ? 'center' : shape.align === 'r' ? 'right' : 'left';
        const frame = shape.textFrame;
        if (frame) {
          prompt.style.padding = [
            frame.margins.top,
            frame.margins.right,
            frame.margins.bottom,
            frame.margins.left,
          ]
            .map((margin) => margin * ratio() + 'px')
            .join(' ');
          prompt.style.alignContent =
            frame.anchor === 'center' ? 'center' : frame.anchor === 'bottom' ? 'end' : 'start';
        }
        hit.append(prompt);
      }
      hit.onpointerdown = (event) => beginDrag(event, shape);
      hit.ondblclick = (event) => {
        event.stopPropagation();
        if (shape.table) startCellAt(shape, event);
        else if (shape.textable) startText(shape);
      };
      if (selected.has(shape.id))
        for (const handle of shape.kind === 'connector'
          ? ['start', 'end']
          : [
              'nw',
              'n',
              'ne',
              'e',
              'se',
              's',
              'sw',
              'w',
              ...(cropping === shape.id ? [] : ['rotate']),
            ]) {
          const node = document.createElement('span');
          node.className = 'shape-handle' + (cropping === shape.id ? ' crop-handle' : '');
          node.dataset.handle = handle;
          node.setAttribute(
            'aria-label',
            handle === 'rotate' ? 'Rotate' : (cropping === shape.id ? 'Crop ' : 'Resize ') + handle,
          );
          if (shape.kind === 'connector') {
            node.setAttribute('aria-label', handle === 'start' ? 'Start point' : 'End point');
            positionEndpoint(node, shape.bounds, shape.flip, handle);
          }
          hit.append(node);
        }
      if (cropping === shape.id)
        renderPictureCropOverlay(hit, shape, elementFor(shape.id)?.querySelector('image') ?? null);
      appendShapeOverlay(shape, hit);
    }
  host.style.cursor = tool ? 'crosshair' : '';
  if (tool) layer.style.pointerEvents = 'none';
  if (currentRevision !== state.revision) {
    currentRevision = state.revision;
    if (document.activeElement !== byId('speaker-notes'))
      byId<HTMLTextAreaElement>('speaker-notes').value = slideModel()?.notes ?? '';
  }
  document
    .querySelectorAll('.thumbnail')
    .forEach((item, index) =>
      item.classList.toggle('slide-hidden', !!model()?.slides[index]?.hidden),
    );
  renderSections();
  renderSelection();
  updateControls();
}
function renameSelectionShape(shape: EditorShape, button: HTMLButtonElement) {
  if (busy || !button.isConnected) return;
  const state = office.getState();
  const key = slideModel()?.key;
  const input = document.createElement('input');
  input.value = shape.name;
  input.setAttribute('aria-label', 'Object name');
  input.style.cssText = 'min-width:0;flex:1';
  let finished = false;
  const finish = async (cancel: boolean) => {
    if (finished) return;
    finished = true;
    const connected = input.isConnected;
    const name = input.value;
    input.disabled = true;
    if (!cancel && connected && slideModel()?.key === key && name !== shape.name) {
      await send(
        { type: 'update', slide: state.index, ids: [shape.id], changes: { name } },
        undefined,
        state.revision,
      );
    }
    if (slideModel()?.key === key) renderSelection();
  };
  input.onkeydown = (event) => {
    event.stopPropagation();
    if (event.key === 'Enter' || event.key === 'Escape') {
      event.preventDefault();
      void finish(event.key === 'Escape');
    }
  };
  input.onblur = () => void finish(false);
  button.replaceWith(input);
  input.focus();
  input.select();
}
let stopSelectionDragScroll = () => {};
function renderSelection() {
  stopSelectionDragScroll();
  const pane = byId('selection-pane');
  pane.replaceChildren();
  const title = document.createElement('h2');
  title.textContent = 'Selection Pane';
  const header = document.createElement('header');
  header.append(title);
  const close = document.createElement('button');
  close.textContent = '×';
  close.setAttribute('aria-label', 'Close Selection Pane');
  close.onclick = () => {
    stopSelectionDragScroll();
    pane.hidden = true;
  };
  header.append(close);
  pane.append(header);
  const visibilityActions = document.createElement('div');
  visibilityActions.style.cssText = 'display:flex;gap:6px;margin:6px 0';
  for (const [label, hidden] of [
    ['Show All', false],
    ['Hide All', true],
  ] as const) {
    const action = document.createElement('button');
    action.textContent = label;
    action.disabled = busy || !allShapes().length;
    action.onclick = () => {
      if (busy || !action.isConnected) return;
      const state = office.getState();
      const ids = allShapes()
        .filter((shape) => !!shape.hidden !== hidden)
        .map((shape) => shape.id);
      if (!ids.length) return;
      void send(
        { type: 'update', slide: state.index, ids, changes: { hidden } },
        undefined,
        state.revision,
      );
    };
    visibilityActions.append(action);
  }
  pane.append(visibilityActions);
  let dragged: { ids: number[]; siblings: number[]; revision: number; slide: number } | null = null;
  let scrollFrame = 0;
  let scrollVelocity = 0;
  let previousFrame = 0;
  stopSelectionDragScroll = () => {
    cancelAnimationFrame(scrollFrame);
    scrollFrame = 0;
    scrollVelocity = 0;
    previousFrame = 0;
  };
  const scrollDuringDrag = (time: number) => {
    if (!dragged || pane.hidden || !pane.isConnected || !scrollVelocity) {
      stopSelectionDragScroll();
      return;
    }
    const elapsed = previousFrame ? Math.min(time - previousFrame, 50) : 16;
    previousFrame = time;
    pane.scrollTop += (scrollVelocity * elapsed) / 1000;
    scrollFrame = requestAnimationFrame(scrollDuringDrag);
  };
  pane.ondragover = (event) => {
    if (!dragged) return;
    const bounds = pane.getBoundingClientRect();
    const edge = Math.min(40, bounds.height / 4);
    const distance = event.clientY - bounds.top;
    scrollVelocity =
      distance < edge
        ? -500 * Math.max(0, 1 - distance / edge)
        : distance > bounds.height - edge
          ? 500 * Math.max(0, 1 - (bounds.height - distance) / edge)
          : 0;
    if (!scrollVelocity) stopSelectionDragScroll();
    else if (!scrollFrame) scrollFrame = requestAnimationFrame(scrollDuringDrag);
  };
  pane.ondragleave = (event) => {
    if (!(event.relatedTarget instanceof Node) || !pane.contains(event.relatedTarget))
      stopSelectionDragScroll();
  };
  pane.ondrop = () => {
    dragged = null;
    stopSelectionDragScroll();
  };
  const addShapes = (shapes: EditorShape[], depth = 0) => {
    for (const shape of [...shapes].reverse()) {
      const row = document.createElement('div');
      row.style.cssText = `display:flex;align-items:center;margin-left:${depth * 14}px`;
      if (shape.children?.length) {
        const expand = document.createElement('button');
        const expanded = expandedGroups.has(shape.id);
        expand.textContent = expanded ? '▾' : '▸';
        expand.style.cssText = 'flex:0 0 22px;width:22px';
        expand.setAttribute('aria-label', `${expanded ? 'Collapse' : 'Expand'} ${shape.name}`);
        expand.setAttribute('aria-expanded', String(expanded));
        expand.onclick = () => {
          if (expanded) expandedGroups.delete(shape.id);
          else expandedGroups.add(shape.id);
          renderSelection();
        };
        row.append(expand);
      }
      const button = document.createElement('button');
      button.textContent = shape.name;
      button.title = shape.name;
      button.dataset.selectionId = String(shape.id);
      button.draggable = true;
      button.ondragstart = (event) => {
        if (busy) {
          event.preventDefault();
          return;
        }
        const state = office.getState();
        const siblings = shapes.map((item) => item.id);
        const ids = selected.has(shape.id) ? siblings.filter((id) => selected.has(id)) : [shape.id];
        dragged = { ids, siblings, revision: state.revision, slide: state.index };
        event.dataTransfer?.setData('text/plain', shape.name);
        if (event.dataTransfer) event.dataTransfer.effectAllowed = 'move';
      };
      button.ondragend = () => {
        stopSelectionDragScroll();
        dragged = null;
        pane
          .querySelectorAll<HTMLElement>('[data-selection-id]')
          .forEach((item) => (item.style.boxShadow = ''));
      };
      button.ondragover = (event) => {
        if (!dragged || !dragged.siblings.includes(shape.id) || dragged.ids.includes(shape.id))
          return;
        event.preventDefault();
        if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';
        const above = event.clientY < button.getBoundingClientRect().top + button.offsetHeight / 2;
        button.style.boxShadow = `inset 0 ${above ? 2 : -2}px var(--office-accent, #d35230)`;
      };
      button.ondragleave = () => {
        button.style.boxShadow = '';
      };
      button.ondrop = (event) => {
        event.preventDefault();
        button.style.boxShadow = '';
        const moving = dragged;
        dragged = null;
        if (busy || !moving || !moving.siblings.includes(shape.id) || moving.ids.includes(shape.id))
          return;
        const above = event.clientY < button.getBoundingClientRect().top + button.offsetHeight / 2;
        void send(
          {
            type: 'order',
            slide: moving.slide,
            ids: moving.ids,
            orderPlacement: { target: shape.id, side: above ? 'after' : 'before' },
          },
          undefined,
          moving.revision,
        );
      };

      button.setAttribute('aria-pressed', String(selected.has(shape.id)));
      button.disabled = shape.parentTransform === null;
      button.onclick = (event) => {
        if (event.detail === 2 && !event.shiftKey && !event.metaKey && !event.ctrlKey) {
          renameSelectionShape(shape, button);
          return;
        }
        selectedAnimation = null;
        const extend = event.shiftKey || event.metaKey || event.ctrlKey;
        const remove = extend && selected.has(shape.id);
        if (!extend) selected.clear();
        // A group and its descendants cannot participate in the same selection.
        const descendants = new Set(flattenShapes(shape.children ?? []).map((s) => s.id));
        for (const other of allShapes()) {
          if (
            descendants.has(other.id) ||
            flattenShapes(other.children ?? []).some((s) => s.id === shape.id)
          )
            selected.delete(other.id);
        }
        if (remove) selected.delete(shape.id);
        else selected.add(shape.id);
        render();
        focusCanvas();
      };
      button.onkeydown = (event) => {
        if (event.key !== 'F2') return;
        event.preventDefault();
        event.stopPropagation();
        renameSelectionShape(shape, button);
      };
      button.style.flex = '1';
      row.append(button);
      const visibility = document.createElement('button');
      visibility.setAttribute('aria-label', `${shape.hidden ? 'Show' : 'Hide'} ${shape.name}`);
      visibility.title = shape.hidden ? 'Show' : 'Hide';
      visibility.style.cssText = 'flex:0 0 28px;width:28px';
      visibility.innerHTML = shape.hidden
        ? '<svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true"><path d="M2 2L14 14M1 8Q8 -1 15 8Q8 17 1 8Z" fill="none" stroke="currentColor"/></svg>'
        : '<svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true"><path d="M1 8Q8 -1 15 8Q8 17 1 8Z" fill="none" stroke="currentColor"/><circle cx="8" cy="8" r="2" fill="currentColor"/></svg>';
      visibility.disabled = busy;
      visibility.onclick = () => {
        const state = office.getState();
        if (busy || !visibility.isConnected) return;
        void send(
          {
            type: 'update',
            slide: state.index,
            ids: [shape.id],
            changes: { hidden: !shape.hidden },
          },
          undefined,
          state.revision,
        );
      };
      row.append(visibility);
      pane.append(row);
      if (shape.children && expandedGroups.has(shape.id)) addShapes(shape.children, depth + 1);
    }
  };
  addShapes(slideModel()?.shapes ?? []);
}
function measuredTextBounds(field: RichTextField, shape: EditorShape, scale: number) {
  const size = measureRichTextBox(field);
  const wrapped = shape.textFrame?.wrap !== false;
  const vertical = shape.textDirection !== 'horz';
  return resizedTextBounds(
    shape,
    wrapped && !vertical ? shape.bounds!.w : size.width / scale,
    wrapped && vertical ? shape.bounds!.h : size.height / scale,
  );
}
function styleTextField(field: RichTextField, shape: EditorShape, scale: number) {
  const frame = shape.textFrame;
  Object.assign(field.style, {
    fontFamily: textFontFamily(shape.format.font),
    fontSize:
      (shape.format.size ?? 18) * 12700 * scale * (shape.autoFitParams?.fontScale ?? 1) + 'px',
    fontWeight: shape.format.bold ? 'bold' : 'normal',
    fontStyle: shape.format.italic ? 'italic' : 'normal',
    color: color(shape.format.color, '#000000'),
    background: color(shape.fill, '#ffffff'),
    transform: `rotate(${shape.rotation + (shape.textDirection === 'vert270' ? 180 : 0)}deg)`,
    writingMode:
      shape.textDirection === 'horz'
        ? 'horizontal-tb'
        : shape.textDirection === 'vert270' || shape.textDirection === 'mongolianVert'
          ? 'vertical-lr'
          : 'vertical-rl',
    textOrientation:
      shape.textDirection === 'wordArtVert' || shape.textDirection === 'wordArtVertRtl'
        ? 'upright'
        : shape.textDirection === 'eaVert'
          ? 'mixed'
          : 'sideways',
    textAlign: shape.align === 'ctr' ? 'center' : shape.align === 'r' ? 'right' : 'left',
    padding: frame
      ? (shape.textDirection === 'vert270'
          ? [frame.margins.bottom, frame.margins.left, frame.margins.top, frame.margins.right]
          : [frame.margins.top, frame.margins.right, frame.margins.bottom, frame.margins.left]
        )
          .map((margin) => margin * scale + 'px')
          .join(' ')
      : '0',
    whiteSpace: frame?.wrap === false ? 'pre' : 'pre-wrap',
    alignContent:
      frame?.anchor === 'center' ? 'center' : frame?.anchor === 'bottom' ? 'end' : 'start',
    columnCount: String(frame?.columns?.count ?? 1),
    columnGap: (frame?.columns?.gapEmu ?? 0) * scale + 'px',
    columnFill: 'auto',
  });
}
function startCellAt(
  shape: EditorShape,
  event: MouseEvent,
  origin = event.currentTarget as HTMLElement,
) {
  const rect = origin.getBoundingClientRect();
  const scale = ratio(),
    box = shape.bounds!;
  const angle = (-shape.rotation * Math.PI) / 180;
  const { x: dx, y: dy } = localVector(
    shape,
    (event.clientX - (rect.left + rect.width / 2)) / scale,
    (event.clientY - (rect.top + rect.height / 2)) / scale,
  );
  const x = dx * Math.cos(angle) - dy * Math.sin(angle) + box.w / 2;
  const y = dx * Math.sin(angle) + dy * Math.cos(angle) + box.h / 2;
  const find = (values: number[], point: number, size: number) => {
    const total = values.reduce((a, b) => a + b, 0);
    let end = 0;
    return values.findIndex((value) => (end += (value * size) / total) > point);
  };
  const row = find(shape.table!.rowHeights, y, box.h);
  const column = find(shape.table!.columnWidths, x, box.w);
  for (let r = 0; r <= row; r++)
    for (let c = 0; c <= column; c++) {
      const cell = shape.table!.cells[r]?.[c];
      if (
        cell &&
        !cell.span.hMerge &&
        !cell.span.vMerge &&
        r + cell.span.rowSpan > row &&
        c + cell.span.gridSpan > column
      ) {
        startCell(shape, r, c);
        return;
      }
    }
}
function startCell(
  table: EditorShape,
  row: number,
  column: number,
  range?: { start: number; end: number },
) {
  const rendered = elementFor(table.id)?.querySelector<SVGRectElement>(
    `[data-pptx-cell="${row},${column}"]`,
  );
  // Use the renderer's resolved appearance so editing follows its table style.
  // A transparent cell sits on the renderer's white table backdrop.
  const shape = tableCellShape(
    table,
    row,
    column,
    rendered
      ? {
          fill: color(rendered.getAttribute('fill'), '#ffffff'),
          color: rendered.dataset.pptxTextColor ?? '#000000',
          font: rendered.dataset.pptxFont ?? 'Calibri',
        }
      : undefined,
  );
  if (shape) startText(shape, range, { row, column });
}
function startText(
  shape: EditorShape,
  range?: { start: number; end: number },
  cell?: { row: number; column: number },
) {
  if (busy || office.getState().building || !shape.bounds || !shape.textable || textEdit) return;
  cellSelection = null;
  lastClick = null;
  selected.clear();
  selected.add(shape.id);
  render();
  const field = createRichTextField();
  field.className = 'direct-text';
  field.setAttribute(
    'aria-label',
    cell ? `Edit cell ${cell.row + 1}, ${cell.column + 1}` : 'Edit text',
  );
  field.value = shape.text;
  setBox(field, shape.bounds);
  const scale = ratio();
  styleTextField(field, shape, scale);
  textEdit = {
    ...(cell ? { cell } : {}),
    field,
    undoText: () => false,
    undoCount: 0,
    redoCount: 0,
    shape,
    value: shape.text,
    edits: [],
    runs: shape.runs.map((run) => ({ ...run })),
    paragraphs: shape.paragraphs.map((paragraph) => ({ ...paragraph })),
    autoFitParams: shape.autoFitParams,
    bounds: shape.bounds,
    insertion: null,
    revision: office.getState().revision,
    slide: office.getState().index,
  };
  appendShapeOverlay(shape, field);
  field.focus();
  const renderText = (recalculate = false) => {
    if (!textEdit || textEdit.field !== field) return;
    if (shape.autoFit === 'normal' && recalculate) {
      textEdit.autoFitParams = fitRichText(
        field,
        textEdit.runs,
        scale,
        textEdit.paragraphs,
        shape.format.size ?? 18,
        shape.autoFitParams?.lnSpcReduction ?? 0,
      );
    } else {
      field.style.fontSize =
        (shape.format.size ?? 18) * 12700 * scale * (textEdit.autoFitParams?.fontScale ?? 1) + 'px';
      field.renderRuns(textEdit.runs, scale, textEdit.paragraphs, textEdit.autoFitParams);
    }
    if (shape.autoFit === 'shape' && recalculate) {
      textEdit.bounds = measuredTextBounds(field, shape, scale);
    }
    if (textEdit.bounds) setBox(field, textEdit.bounds);
    centerRichText(field, shape.anchorCenter);
  };
  renderText(shape.autoFitParams?.fontScale === 1);
  if (range) field.setSelectionRange(range.start, range.end);
  else field.select();
  const capture = () => ({
    value: textEdit!.value,
    runs: textEdit!.runs.map((run) => ({ ...run })),
    paragraphs: textEdit!.paragraphs.map((paragraph) => ({ ...paragraph })),
    edits: [...textEdit!.edits],
    autoFitParams: textEdit!.autoFitParams,
    bounds: textEdit!.bounds,
    start: field.selectionStart,
    end: field.selectionEnd,
  });
  const textHistory = [capture()];
  let textCursor = 0;
  let compositionCursor: number | null = null;
  field.addEventListener('compositionstart', () => {
    compositionCursor = textCursor;
  });
  textEdit.undoText = (redo) => {
    const next = textCursor + (redo ? 1 : -1);
    if (!textEdit || next < 0 || next >= textHistory.length) return false;
    textCursor = next;
    const snapshot = textHistory[next]!;
    textEdit.value = snapshot.value;
    textEdit.runs = snapshot.runs.map((run) => ({ ...run }));
    textEdit.paragraphs = snapshot.paragraphs.map((paragraph) => ({ ...paragraph }));
    textEdit.edits = [...snapshot.edits];
    textEdit.autoFitParams = snapshot.autoFitParams;
    textEdit.bounds = snapshot.bounds;
    textEdit.insertion = null;
    textEdit.undoCount = next;
    textEdit.redoCount = textHistory.length - next - 1;
    field.value = snapshot.value;
    field.focus();
    field.setSelectionRange(snapshot.start, snapshot.end);
    renderText();
    updateControls();
    return true;
  };
  let inputPosition: { start: number; end: number; type: string } | null = null;
  field.onbeforeinput = (event) => {
    if (compositionCursor === null || compositionCursor === textCursor) {
      textHistory[textCursor]!.start = field.selectionStart;
      textHistory[textCursor]!.end = field.selectionEnd;
    }
    inputPosition = {
      start: field.selectionStart,
      end: field.selectionEnd,
      type: (event as InputEvent).inputType,
    };
  };
  field.oninput = (event) => {
    if (!textEdit) return;
    const before = textEdit.value,
      after = field.value;
    let start = 0,
      end = before.length,
      to = after.length;
    const prefixLimit = inputPosition
      ? Math.max(
          0,
          inputPosition.start -
            (inputPosition.start === inputPosition.end && inputPosition.type.endsWith('Backward')
              ? Math.max(0, before.length - after.length)
              : 0),
        )
      : end;
    while (start < end && start < to && start < prefixLimit && before[start] === after[start])
      start++;
    inputPosition = null;
    // Keep surrogate pairs intact even when two emoji share a high surrogate.
    if (start > 0 && /[\uD800-\uDBFF]/.test(before[start - 1]!)) start--;
    while (end > start && to > start && before[end - 1] === after[to - 1]) {
      end--;
      to--;
    }
    if (end < before.length && /[\uDC00-\uDFFF]/.test(before[end]!)) {
      end++;
      to++;
    }
    const inserted = after.slice(start, to);
    const format = textEdit.insertion?.format;
    textEdit.edits.push({
      start,
      end,
      text: inserted,
      ...(format ? { format: { ...format } } : {}),
    });
    const inherited =
      textEdit.runs.find(
        (r) =>
          r.start <= (end > start ? start : Math.max(0, start - 1)) &&
          (end > start ? start : Math.max(0, start - 1)) < r.end,
      )?.format ??
      textEdit.runs[0]?.format ??
      textEdit.shape.format;
    const runs: EditorShape['runs'] = [];
    for (const run of textEdit.runs)
      if (run.start < start) runs.push({ ...run, end: Math.min(start, run.end) });
    if (inserted)
      runs.push({ start, end: start + inserted.length, format: { ...inherited, ...format } });
    const shift = inserted.length - (end - start);
    for (const run of textEdit.runs)
      if (run.end > end)
        runs.push({ ...run, start: Math.max(end, run.start) + shift, end: run.end + shift });
    textEdit.runs = runs;
    textEdit.paragraphs = replaceParagraphRange(textEdit.paragraphs, start, end, inserted);
    textEdit.value = after;
    if (!(event as InputEvent).isComposing) renderText(true);
    if (textEdit.insertion) textEdit.insertion.position = field.selectionStart;
    if (compositionCursor !== null) textCursor = compositionCursor;
    textHistory.splice(textCursor + 1);
    textHistory.push(capture());
    textCursor++;
    textEdit.undoCount = textCursor;
    textEdit.redoCount = 0;
    updateControls();
  };
  field.addEventListener('compositionend', () => {
    queueMicrotask(() => {
      compositionCursor = null;
      if (textEdit?.field === field) {
        renderText(true);
        textHistory[textCursor]!.autoFitParams = textEdit.autoFitParams;
        textHistory[textCursor]!.bounds = textEdit.bounds;
      }
    });
  });
  field.onselect = () => {
    if (
      textEdit?.insertion &&
      (field.selectionStart !== field.selectionEnd ||
        field.selectionStart !== textEdit.insertion.position)
    )
      textEdit.insertion = null;
    updateControls();
  };
  field.onkeydown = (event) => {
    event.stopPropagation();
    if (textFormatShortcut(event, field)) return;
    if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'z') {
      event.preventDefault();
      void action(event.shiftKey ? 'redo' : 'undo', field);
      return;
    }
    if ((event.metaKey || event.ctrlKey) && ['b', 'i', 'u'].includes(event.key.toLowerCase())) {
      event.preventDefault();
      const name = { b: 'bold', i: 'italic', u: 'underline' }[event.key.toLowerCase()]!;
      void action(name, document.querySelector(`[data-edit="${name}"]`)!);
      return;
    }
    if (cell && event.key === 'Tab') {
      event.preventDefault();
      const positions = shape.table!.cells.flatMap((row, r) =>
        row.flatMap((entry, c) =>
          entry.span.hMerge || entry.span.vMerge ? [] : [{ row: r, column: c }],
        ),
      );
      const index = positions.findIndex((p) => p.row === cell.row && p.column === cell.column);
      const next = positions[index + (event.shiftKey ? -1 : 1)];
      const slide = office.getState().index;
      const slideKey = slideModel()?.key;
      const append = !event.shiftKey && index === positions.length - 1;
      void finishText().then(async (ok) => {
        if (!ok || slideModel()?.key !== slideKey || office.getState().index !== slide) return;
        if (append) {
          if (!(await send({ type: 'table-row-append', slide, ids: [shape.id] }))) return;
          if (slideModel()?.key !== slideKey || office.getState().index !== slide) return;
        }
        const updated = allShapes().find((s) => s.id === shape.id);
        const destination =
          append && updated?.table ? { row: updated.table.cells.length - 1, column: 0 } : next;
        if (updated && destination) startCell(updated, destination.row, destination.column);
      });
      return;
    }
    if (event.key === 'Escape') {
      event.preventDefault();
      void finishText();
    }
    if ((event.metaKey || event.ctrlKey) && event.key === 'Enter') {
      event.preventDefault();
      void finishText();
    }
  };
  field.onblur = (event) => {
    const target = event.relatedTarget;
    if (
      !(
        target instanceof Element &&
        target.closest(
          '.office-header, .office-menu, .find-replace-dialog, .table-margins-dialog, .font-dialog, .paragraph-dialog',
        )
      )
    )
      void finishText();
  };
}
async function finishText() {
  const edit = textEdit;
  if (!edit) return true;
  textEdit = null;
  if (!edit.edits.length) {
    edit.field.dispose();
    edit.field.remove();
    render();
    return true;
  }
  const success = await send(
    {
      type: edit.cell ? 'table-cell-text' : 'update',
      slide: edit.slide,
      ids: [edit.shape.id],
      ...(edit.cell ? { cell: { ...edit.cell, edits: edit.edits } } : {}),
      changes: {
        textEdits: edit.edits,
        ...(edit.shape.autoFit === 'normal' && edit.autoFitParams
          ? { autoFitParams: edit.autoFitParams }
          : {}),
        ...(edit.shape.autoFit === 'shape' && edit.bounds ? { bounds: edit.bounds } : {}),
      },
    },
    undefined,
    edit.revision,
  );
  if (!success) {
    textEdit = edit;
    if (!shadow.contains(edit.field)) layer.append(edit.field);
    edit.field.focus();
    return false;
  }
  edit.field.dispose();
  edit.field.remove();
  render();
  focusCanvas();
  return true;
}
function connectionAt(p: { x: number; y: number }, exclude?: number, disabled = false) {
  layer.querySelectorAll('.connection-site').forEach((node) => node.remove());
  let best: {
    point: { x: number; y: number };
    connection: NonNullable<EditorShape['connections']['start']>;
  } | null = null;
  let distance = 9;
  if (disabled) return null;
  for (const shape of slideModel()?.shapes ?? []) {
    if (shape.id === exclude || !shape.bounds || !shape.connectionSites.length) continue;
    const sites = shape.connectionSites;
    const near = sites.some((site) => Math.hypot(site.x - p.x, site.y - p.y) * ratio() < 40);
    if (!near) continue;
    for (const [siteIndex, site] of sites.entries()) {
      const d = Math.hypot(site.x - p.x, site.y - p.y) * ratio();
      const dot = document.createElement('span');
      dot.className = 'connection-site';
      dot.style.cssText = `position:absolute;pointer-events:none;left:${site.x * ratio() - 3}px;top:${site.y * ratio() - 3}px;width:6px;height:6px;border:1px solid #a33165;border-radius:50%;background:${d < 9 ? '#a33165' : '#fff'};box-sizing:border-box`;
      layer.append(dot);
      if (d < distance) {
        best = { point: site, connection: { shapeId: shape.id, siteIndex } };
        distance = d;
      }
    }
  }
  return best;
}
function positionEndpoint(
  node: HTMLElement,
  box: NonNullable<EditorShape['bounds']>,
  flip: EditorShape['flip'],
  handle: string,
) {
  const end = handle === 'end';
  const w = box.w * ratio(),
    h = box.h * ratio();
  node.style.left = `${(end !== !!flip?.horizontal ? w : 0) + (w < 8 ? (8 - w) / 2 : 0)}px`;
  node.style.top = `${(end !== !!flip?.vertical ? h : 0) + (h < 8 ? (8 - h) / 2 : 0)}px`;
  node.style.borderRadius = '50%';
  node.style.cursor = 'crosshair';
}
interface Drag {
  pointer: number;
  slide: number;
  startX: number;
  startY: number;
  shapes: EditorShape[];
  handle: string;
  target: number;
  editPlaceholder?: boolean;
  moved: boolean;
  boxes: Map<number, NonNullable<EditorShape['bounds']>>;
  rotation?: number;
  crop?: ReturnType<typeof dragPictureCrop>;
  cropping?: boolean;
  endpoint?: {
    bounds: NonNullable<EditorShape['bounds']>;
    flip: NonNullable<EditorShape['flip']>;
    connections: NonNullable<EditCommand['connections']>;
  };
  revision: number;
  previewSequence?: number;
}
let drag: Drag | null = null;
let lastClick: { id: number; time: number } | null = null;
let dragPreviewTask: Promise<void> | null = null;
let dragPreviewTimer: ReturnType<typeof setTimeout> | undefined;
let dragPreviewOriginal: { node: SVGSVGElement; visibility: string } | null = null;
function dragCommand(session: Drag): EditCommand {
  return {
    type: 'update',
    slide: session.slide,
    ids: session.shapes.map((shape) => shape.id),
    ...(session.handle && session.handle !== 'rotate'
      ? { tableResizeAnchor: session.handle.includes('n') ? ('bottom' as const) : ('top' as const) }
      : {}),
    ...(session.crop
      ? { changes: session.crop }
      : session.endpoint
        ? { changes: session.endpoint }
        : session.rotation !== undefined
          ? { changes: { rotation: session.rotation } }
          : { positions: [...session.boxes].map(([id, bounds]) => ({ id, bounds })) }),
  };
}
function clearDragPreview() {
  clearTimeout(dragPreviewTimer);
  dragPreviewTimer = undefined;
  layer.querySelector('.drag-preview')?.remove();
  if (dragPreviewOriginal)
    dragPreviewOriginal.node.style.visibility = dragPreviewOriginal.visibility;
  dragPreviewOriginal = null;
}
function scheduleDragPreview() {
  if (!drag || dragPreviewTask || dragPreviewTimer !== undefined) return;
  dragPreviewTimer = setTimeout(() => {
    dragPreviewTimer = undefined;
    const session = drag;
    if (!session) return;
    const sequence = session.previewSequence;
    const command = dragCommand(session);
    dragPreviewTask = (async () => {
      try {
        const response = await fetch('/edit', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            revision: session.revision,
            command,
            preview: true,
            previewRender: true,
          }),
        });
        if (!response.ok) return;
        const result = await response.json();
        if (
          drag !== session ||
          office.getState().index !== session.slide ||
          session.previewSequence !== sequence ||
          office.getState().revision !== session.revision ||
          !result.svg
        )
          return;
        const original = [...shadow.children].find(
          (node): node is SVGSVGElement => node instanceof SVGSVGElement,
        );
        if (!original) return;
        let preview = layer.querySelector<HTMLDivElement>('.drag-preview');
        if (!preview) {
          preview = document.createElement('div');
          preview.className = 'drag-preview';
          preview.style.cssText = 'position:absolute;inset:0;pointer-events:none';
          layer.prepend(preview);
        }
        preview.innerHTML = result.svg;
        const svg = preview.querySelector('svg');
        // Keep temporary paint servers and clipping paths distinct from the saved slide.
        const ids = new Map(
          [...preview.querySelectorAll('[id]')].map((node) => [node.id, 'drag-preview-' + node.id]),
        );
        for (const node of preview.querySelectorAll('*')) {
          for (const attribute of node.attributes) {
            let value = attribute.value.replace(/url\(#([^)]*)\)/g, (whole, id) =>
              ids.has(id) ? `url(#${ids.get(id)})` : whole,
            );
            if (attribute.localName === 'href' && value.startsWith('#') && ids.has(value.slice(1)))
              value = '#' + ids.get(value.slice(1));
            if (attribute.name === 'id') value = ids.get(value) ?? value;
            if (value !== attribute.value)
              node.setAttributeNS(attribute.namespaceURI, attribute.name, value);
          }
        }
        if (svg) {
          svg.style.width = '100%';
          svg.style.height = '100%';
        }
        if (!dragPreviewOriginal)
          dragPreviewOriginal = { node: original, visibility: original.style.visibility };
        original.style.visibility = 'hidden';
        layer
          .querySelectorAll('.attached-preview,.endpoint-preview')
          .forEach((node) => node.remove());
      } catch {
        // A transient render failure must not prevent committing the edit.
      } finally {
        dragPreviewTask = null;
        if (drag && (drag !== session || drag.previewSequence !== sequence)) scheduleDragPreview();
      }
    })();
  }, 80);
}

function beginCellSelection(event: PointerEvent, shape: EditorShape): boolean {
  const rect = (event.currentTarget as HTMLElement).getBoundingClientRect();
  const box = shape.bounds!,
    table = shape.table!,
    scale = ratio();
  const angle = (-shape.rotation * Math.PI) / 180;
  const local = (e: PointerEvent) => {
    const { x: dx, y: dy } = localVector(
      shape,
      (e.clientX - rect.left - rect.width / 2) / scale,
      (e.clientY - rect.top - rect.height / 2) / scale,
    );
    return {
      x: dx * Math.cos(angle) - dy * Math.sin(angle) + box.w / 2,
      y: dx * Math.sin(angle) + dy * Math.cos(angle) + box.h / 2,
    };
  };
  const initial = local(event);
  // The outer frame remains available for moving the entire table.
  if (Math.min(initial.x, initial.y, box.w - initial.x, box.h - initial.y) * scale < 5)
    return false;
  const indexAt = (sizes: number[], position: number, size: number) => {
    const total = sizes.reduce((a, b) => a + b, 0);
    let edge = 0;
    const index = sizes.findIndex((value) => (edge += (value / total) * size) > position);
    return index < 0 ? sizes.length - 1 : index;
  };
  const cellAt = (e: PointerEvent) => {
    const p = local(e);
    return {
      row: indexAt(table.rowHeights, p.y, box.h),
      column: indexAt(table.columnWidths, p.x, box.w),
    };
  };
  const first = cellAt(event);
  const state = office.getState();
  let moved = false;
  const update = (e: PointerEvent) => {
    const last = cellAt(e);
    const row = Math.min(first.row, last.row),
      column = Math.min(first.column, last.column);
    const bottom = Math.max(first.row, last.row) + 1,
      right = Math.max(first.column, last.column) + 1;
    cellSelection = {
      id: shape.id,
      slide: state.index,
      revision: state.revision,
      range: expandTableSelection(table, {
        row,
        column,
        rows: bottom - row,
        columns: right - column,
      }),
    };
    render();
  };
  const controller = new AbortController();
  const options = { signal: controller.signal };
  host.addEventListener(
    'pointermove',
    (e) => {
      if (e.pointerId !== event.pointerId) return;
      if (Math.hypot(e.clientX - event.clientX, e.clientY - event.clientY) >= 3) moved = true;
      if (moved) update(e);
    },
    options,
  );
  const finish = (e: PointerEvent) => {
    if (e.pointerId !== event.pointerId) return;
    controller.abort();
    if (host.hasPointerCapture(event.pointerId)) host.releasePointerCapture(event.pointerId);
    if (
      e.type !== 'pointerup' ||
      office.getState().revision !== state.revision ||
      office.getState().index !== state.index
    ) {
      cellSelection = null;
      render();
      return;
    }
    update(e);
    if (!moved && cellSelection)
      startCell(shape, cellSelection.range.row, cellSelection.range.column);
    else focusCanvas();
  };
  host.addEventListener('pointerup', finish, options);
  host.addEventListener('pointercancel', finish, options);
  host.addEventListener('lostpointercapture', finish, options);
  host.setPointerCapture(event.pointerId);
  lastClick = null;
  return true;
}
function beginDrag(event: PointerEvent, shape: EditorShape) {
  if (tool || busy || event.button !== 0) return;
  selectedAnimation = null;
  event.preventDefault();
  event.stopPropagation();
  if (textEdit) {
    void finishText();
    return;
  }
  if (
    shape.children?.length &&
    selected.has(shape.id) &&
    !(event.target as HTMLElement).dataset.handle &&
    !event.shiftKey &&
    !event.metaKey &&
    !event.ctrlKey
  ) {
    const p = point(event);
    const child = [...shape.children].reverse().find((item) => {
      if (item.hidden || !item.bounds || item.parentTransform === null) return false;
      const m = item.parentTransform;
      const local = localVector(item, p.x - (m?.[4] ?? 0), p.y - (m?.[5] ?? 0));
      const box = item.bounds,
        angle = (-item.rotation * Math.PI) / 180;
      const dx = local.x - box.x - box.w / 2,
        dy = local.y - box.y - box.h / 2;
      return (
        Math.abs(dx * Math.cos(angle) - dy * Math.sin(angle)) <= box.w / 2 &&
        Math.abs(dx * Math.sin(angle) + dy * Math.cos(angle)) <= box.h / 2
      );
    });
    if (child) {
      expandedGroups.add(shape.id);
      selected.clear();
      beginDrag(event, child);
      return;
    }
  }
  if (
    shape.table &&
    selected.size === 1 &&
    selected.has(shape.id) &&
    !(event.target as HTMLElement).dataset.handle &&
    !event.shiftKey &&
    !event.metaKey &&
    !event.ctrlKey &&
    beginCellSelection(event, shape)
  )
    return;
  cellSelection = null;
  const now = performance.now();
  const doubleClick = lastClick?.id === shape.id && now - lastClick.time < 400;
  lastClick = { id: shape.id, time: now };
  if (
    doubleClick &&
    (shape.textable || shape.table) &&
    !(event.target as HTMLElement).dataset.handle &&
    !event.shiftKey &&
    !event.metaKey &&
    !event.ctrlKey
  ) {
    lastClick = null;
    if (shape.table) startCellAt(shape, event);
    else startText(shape);
    return;
  }
  if (event.shiftKey || event.metaKey || event.ctrlKey) {
    if (selected.has(shape.id)) selected.delete(shape.id);
    else selected.add(shape.id);
  } else if (!selected.has(shape.id)) {
    selected.clear();
    selected.add(shape.id);
  }
  if (selected.has(shape.id)) {
    for (const other of allShapes()) {
      if (
        flattenShapes(shape.children ?? []).some((s) => s.id === other.id) ||
        flattenShapes(other.children ?? []).some((s) => s.id === shape.id)
      )
        selected.delete(other.id);
    }
  }
  const handle = (event.target as HTMLElement).dataset.handle ?? '';
  if (handle) {
    selected.clear();
    selected.add(shape.id);
  }
  render();
  focusCanvas();
  drag = {
    slide: office.getState().index,
    pointer: event.pointerId,
    startX: event.clientX,
    startY: event.clientY,
    shapes: selection(),
    cropping: cropping === shape.id,
    handle,
    target: shape.id,
    editPlaceholder:
      !!shape.placeholder &&
      !shape.text &&
      !handle &&
      !event.shiftKey &&
      !event.metaKey &&
      !event.ctrlKey,
    moved: false,
    boxes: new Map(),
    revision: office.getState().revision,
  };
  host.setPointerCapture(event.pointerId);
}
function previewConnector(shape: EditorShape, frame: ConnectorFrame, className: string) {
  const preview = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  preview.classList.add(className);
  const box = frame.bounds!,
    flip = frame.flip ?? { horizontal: false, vertical: false };
  const w = box.w * ratio(),
    h = box.h * ratio();
  preview.style.cssText = `position:absolute;overflow:visible;pointer-events:none;left:${box.x * ratio()}px;top:${box.y * ratio()}px;width:${Math.max(1, w)}px;height:${Math.max(1, h)}px;transform-origin:${w / 2}px ${h / 2}px;transform:rotate(${frame.rotation}deg)`;
  const path = document.createElementNS(preview.namespaceURI, 'path');
  path.setAttribute(
    'd',
    getPresetShapePath(
      shape.preset === 'line' ? 'straightConnector1' : (shape.preset ?? 'straightConnector1'),
      w,
      h,
    ) ?? `M0 0 L${w} ${h}`,
  );
  path.setAttribute(
    'transform',
    `translate(${flip.horizontal ? w : 0} ${flip.vertical ? h : 0}) scale(${flip.horizontal ? -1 : 1} ${flip.vertical ? -1 : 1})`,
  );
  path.setAttribute('fill', 'none');
  path.setAttribute('stroke', '#777');
  path.setAttribute('stroke-width', '1');
  preview.append(path);
  appendShapeOverlay(shape, preview);
}
host.addEventListener('pointermove', (event) => {
  if (!drag || drag.pointer !== event.pointerId) return;
  let dx = (event.clientX - drag.startX) / ratio(),
    dy = (event.clientY - drag.startY) / ratio();
  if (Math.hypot(dx, dy) * ratio() < 3 && !drag.moved) return;
  drag.moved = true;
  lastClick = null;
  if (event.shiftKey && !drag.handle) {
    if (Math.abs(dx) > Math.abs(dy)) dy = 0;
    else dx = 0;
  }
  layer.querySelectorAll('.smart-guide').forEach((node) => node.remove());
  if (!drag.cropping && !drag.handle && !event.altKey && !event.metaKey && model()) {
    const snapped = snapMove(
      drag.shapes,
      smartGuideTargets(slideModel()?.shapes ?? [], drag.shapes),
      model()!,
      { x: dx, y: dy },
      5 / ratio(),
      event.shiftKey ? (dx === 0 ? 'y' : 'x') : undefined,
      {
        ...guideSettings,
        guides: drawingGuides,
        grid: guideSettings.snap ? gridSpacing() : undefined,
      },
    );
    dx = snapped.x;
    dy = snapped.y;
    for (const guide of guideSettings.smart ? snapped.guides : []) {
      const line = document.createElement('div');
      line.className = 'smart-guide';
      line.dataset.axis = guide.axis;
      line.setAttribute('aria-hidden', 'true');
      const spacing = guide.kind === 'spacing';
      if (spacing) line.dataset.kind = 'spacing';
      const vertical = spacing ? guide.axis === 'y' : guide.axis === 'x';
      line.style.cssText = `position:absolute;pointer-events:none;z-index:2;left:${(vertical ? guide.value : guide.from) * ratio()}px;top:${(vertical ? guide.from : guide.value) * ratio()}px;width:${vertical ? 0 : (guide.to - guide.from) * ratio()}px;height:${vertical ? (guide.to - guide.from) * ratio() : 0}px;border-${vertical ? 'left' : 'top'}:1px dashed #c43c3c`;
      if (spacing) {
        for (const end of ['start', 'end']) {
          const cap = document.createElement('span');
          cap.style.cssText = `position:absolute;${vertical ? 'left:-3px;width:6px;height:0;border-top' : 'top:-3px;height:6px;width:0;border-left'}:1px solid #c43c3c;${vertical ? (end === 'start' ? 'top' : 'bottom') : end === 'start' ? 'left' : 'right'}:0`;
          line.append(cap);
        }
      }
      layer.append(line);
    }
  }
  const slideDx = dx,
    slideDy = dy;
  for (const shape of drag.shapes) {
    if (!shape.bounds) continue;
    ({ x: dx, y: dy } = localVector(shape, slideDx, slideDy));
    const original = shape.bounds,
      box = { ...original };
    const hit = layer.querySelector<HTMLElement>(`[data-shape-id="${shape.id}"]`)!;
    if (drag.cropping) {
      drag.crop = dragPictureCrop(shape, drag.handle, dx, dy, cropAspectRatio);
      drag.boxes.set(shape.id, drag.crop.bounds);
      setBox(hit, drag.crop.bounds);
      renderPictureCropOverlay(
        hit,
        { ...shape, ...drag.crop },
        elementFor(shape.id)?.querySelector('image') ?? null,
      );
      continue;
    }
    if (drag.handle === 'start' || drag.handle === 'end') {
      const angle = (shape.rotation * Math.PI) / 180;
      const c = Math.cos(angle),
        s = Math.sin(angle);
      const start = {
        x: shape.flip?.horizontal ? original.w : 0,
        y: shape.flip?.vertical ? original.h : 0,
      };
      const end = { x: original.w - start.x, y: original.h - start.y };
      const moving = drag.handle === 'start' ? start : end;
      const fixed = drag.handle === 'start' ? end : start;
      moving.x += dx * c + dy * s;
      moving.y += -dx * s + dy * c;
      const snapped = connectionAt(point(event), shape.id, event.altKey);
      if (snapped) {
        const m = shape.parentTransform;
        const local = localVector(
          shape,
          snapped.point.x - (m?.[4] ?? 0),
          snapped.point.y - (m?.[5] ?? 0),
        );
        const sx = local.x - original.x - original.w / 2;
        const sy = local.y - original.y - original.h / 2;
        moving.x = original.w / 2 + sx * c + sy * s;
        moving.y = original.h / 2 - sx * s + sy * c;
      }
      if (event.shiftKey && !snapped) {
        const length = Math.hypot(moving.x - fixed.x, moving.y - fixed.y);
        const direction =
          (Math.round(Math.atan2(moving.y - fixed.y, moving.x - fixed.x) / (Math.PI / 4)) *
            Math.PI) /
          4;
        moving.x = fixed.x + length * Math.cos(direction);
        moving.y = fixed.y + length * Math.sin(direction);
      }
      start.x = Math.round(start.x);
      start.y = Math.round(start.y);
      end.x = Math.round(end.x);
      end.y = Math.round(end.y);
      box.w = Math.abs(end.x - start.x);
      box.h = Math.abs(end.y - start.y);
      const cx = (start.x + end.x - original.w) / 2;
      const cy = (start.y + end.y - original.h) / 2;
      box.x = original.x + original.w / 2 + cx * c - cy * s - box.w / 2;
      box.y = original.y + original.h / 2 + cx * s + cy * c - box.h / 2;
      const flip = { horizontal: start.x > end.x, vertical: start.y > end.y };
      drag.endpoint = {
        bounds: box,
        flip,
        connections: { [drag.handle]: snapped?.connection ?? null },
      };
      layer.querySelector('.endpoint-preview')?.remove();
      previewConnector(shape, { bounds: box, rotation: shape.rotation, flip }, 'endpoint-preview');

      setBox(hit, box);
      hit.style.width = `${Math.max(8, box.w * ratio())}px`;
      hit.style.height = `${Math.max(8, box.h * ratio())}px`;
      hit.style.marginLeft = `${Math.min(0, (box.w * ratio() - 8) / 2)}px`;
      hit.style.marginTop = `${Math.min(0, (box.h * ratio() - 8) / 2)}px`;
      hit
        .querySelectorAll<HTMLElement>('.shape-handle')
        .forEach((node) => positionEndpoint(node, box, flip, node.dataset.handle!));
      continue;
    }
    if (drag.handle === 'rotate') {
      const p = point(event);
      const m = shape.parentTransform;
      const local = localVector(shape, p.x - (m?.[4] ?? 0), p.y - (m?.[5] ?? 0));
      let angle =
        (Math.atan2(local.y - box.y - box.h / 2, local.x - box.x - box.w / 2) * 180) / Math.PI + 90;
      if (event.shiftKey) angle = Math.round(angle / 15) * 15;
      drag.rotation = angle;
      hit.style.transform = `rotate(${angle}deg)`;
      continue;
    }
    if (drag.handle) {
      const radians = (shape.rotation * Math.PI) / 180;
      const localX = dx * Math.cos(radians) + dy * Math.sin(radians),
        localY = -dx * Math.sin(radians) + dy * Math.cos(radians);
      if (drag.handle.includes('e')) box.w = Math.max(12700, original.w + localX);
      if (drag.handle.includes('s')) box.h = Math.max(12700, original.h + localY);
      if (drag.handle.includes('w')) {
        box.w = Math.max(12700, original.w - localX);
      }
      if (drag.handle.includes('n')) {
        box.h = Math.max(12700, original.h - localY);
      }
      if (
        (event.shiftKey || shape.aspectRatioLocked) &&
        drag.handle.length === 2 &&
        original.w > 0 &&
        original.h > 0
      ) {
        // Follow the dominant relative displacement, including a vertical-only drag.
        const widthScale = box.w / original.w;
        const heightScale = box.h / original.h;
        const scale = Math.max(
          12700 / original.w,
          12700 / original.h,
          Math.abs(widthScale - 1) >= Math.abs(heightScale - 1) ? widthScale : heightScale,
        );
        box.w = original.w * scale;
        box.h = original.h * scale;
      }
      const cx = ((box.w - original.w) / 2) * (drag.handle.includes('w') ? -1 : 1),
        cy = ((box.h - original.h) / 2) * (drag.handle.includes('n') ? -1 : 1);
      box.x =
        original.x + original.w / 2 + cx * Math.cos(radians) - cy * Math.sin(radians) - box.w / 2;
      box.y =
        original.y + original.h / 2 + cx * Math.sin(radians) + cy * Math.cos(radians) - box.h / 2;
    } else {
      box.x += dx;
      box.y += dy;
    }
    drag.boxes.set(shape.id, box);
    setBox(hit, box);
    const svg = elementFor(shape.id);
    if (svg && !drag.handle)
      svg.setAttribute(
        'transform',
        `translate(${(box.x - original.x) / 9525} ${(box.y - original.y) / 9525})`,
      );
  }
  drag.previewSequence = (drag.previewSequence ?? 0) + 1;
  scheduleDragPreview();
  layer.querySelectorAll('.attached-preview').forEach((node) => node.remove());
  if (!drag.endpoint) {
    const shapes = slideModel()?.shapes ?? [];
    for (const line of shapes) {
      if (line.kind !== 'connector' || !line.bounds) continue;
      const points = connectorEndpoints({
        ...line,
        bounds: drag.boxes.get(line.id) ?? line.bounds,
      });
      let changed = false;
      for (const end of ['start', 'end'] as const) {
        const attachment = line.connections[end];
        if (!attachment) continue;
        const target = drag.shapes.find((s) => s.id === attachment.shapeId);
        const site = target?.connectionSites[attachment.siteIndex];
        if (!target?.bounds || !site) continue;
        const original = target.bounds;
        const next = drag.boxes.get(target.id) ?? original;
        const angle = (target.rotation * Math.PI) / 180;
        const x = site.x - original.x - original.w / 2;
        const y = site.y - original.y - original.h / 2;
        // Undo the old rotation/flip, then apply the pending size and rotation.
        let localX = x * Math.cos(angle) + y * Math.sin(angle) + original.w / 2;
        let localY = -x * Math.sin(angle) + y * Math.cos(angle) + original.h / 2;
        if (target.flip?.horizontal) localX = original.w - localX;
        if (target.flip?.vertical) localY = original.h - localY;
        points[end] = shapePoint(
          { ...target, bounds: next, rotation: drag.rotation ?? target.rotation },
          {
            x: original.w ? (localX * next.w) / original.w : 0,
            y: original.h ? (localY * next.h) / original.h : 0,
          },
        );
        changed = true;
      }
      if (changed)
        previewConnector(
          line,
          { ...connectorFrame(points.start, points.end, line.rotation), rotation: line.rotation },
          'attached-preview',
        );
    }
  }
});
async function endDrag(event: Pick<PointerEvent, 'pointerId'>, cancel = false) {
  if (!drag || drag.pointer !== event.pointerId) return;
  const finished = drag;
  drag = null;
  if (cancel && host.hasPointerCapture(finished.pointer))
    host.releasePointerCapture(finished.pointer);
  for (const shape of finished.shapes) elementFor(shape.id)?.removeAttribute('transform');
  clearDragPreview();
  if (!cancel && finished.moved) await send(dragCommand(finished), undefined, finished.revision);
  render();
  if (!cancel && !finished.moved && finished.editPlaceholder) {
    const shape = allShapes().find((item) => item.id === finished.target);
    if (shape) startText(shape);
  }
}
document.addEventListener(
  'keydown',
  (event) => {
    if (event.key !== 'Escape' || cropping === null || document.querySelector('dialog[open]'))
      return;
    event.preventDefault();
    event.stopImmediatePropagation();
    if (drag) void endDrag({ pointerId: drag.pointer }, true);
    else {
      cropping = null;
      render();
    }
  },
  true,
);
host.addEventListener('pointerup', (event) => void endDrag(event));
host.addEventListener('pointercancel', (event) => void endDrag(event, true));
host.addEventListener('lostpointercapture', (event) => void endDrag(event, true));
let drawing: {
  x: number;
  y: number;
  pointer: number;
  box: HTMLDivElement;
  preset: string | null;
  selection: number[];
  additive: boolean;
  startConnection?: EditorShape['connections']['start'];
  endConnection?: EditorShape['connections']['end'];
} | null = null;
function linePath(preset: string, w = 24, h = 24) {
  const line = lineTool(preset)!;
  let path = getPresetShapePath(
    line.geometry === 'line' ? 'straightConnector1' : line.geometry,
    w,
    h,
  )!;
  const angle = line.geometry === 'line' ? Math.atan2(h, w) : 0;
  const tip = (x: number, y: number, direction: number) => {
    const dx = Math.cos(direction),
      dy = Math.sin(direction);
    return ` M${x - dx * 7 - dy * 3},${y - dy * 7 + dx * 3} L${x},${y} L${x - dx * 7 + dy * 3},${y - dy * 7 - dx * 3}`;
  };
  if (line.arrows > 0) path += tip(w, h, angle);
  if (line.arrows > 1) path += tip(0, 0, angle + Math.PI);
  return path;
}
function drawingPoint(event: PointerEvent) {
  const p = point(event);
  if (drawing && lineTool(drawing.preset)) {
    const snapped = connectionAt(p, undefined, event.altKey);
    drawing.endConnection = snapped?.connection ?? null;
    if (snapped) return { ...snapped.point };
  }
  if (drawing && event.shiftKey && lineTool(drawing.preset)?.geometry === 'line') {
    const dx = p.x - drawing.x,
      dy = p.y - drawing.y;
    const angle = Math.round(Math.atan2(dy, dx) / (Math.PI / 4)) * (Math.PI / 4);
    const length = Math.hypot(dx, dy);
    p.x = drawing.x + length * Math.cos(angle);
    p.y = drawing.y + length * Math.sin(angle);
    if (Math.abs(p.x - drawing.x) < 1) p.x = drawing.x;
    if (Math.abs(p.y - drawing.y) < 1) p.y = drawing.y;
  }
  return p;
}
function point(event: PointerEvent) {
  const rect = host.getBoundingClientRect();
  return { x: (event.clientX - rect.left) / ratio(), y: (event.clientY - rect.top) / ratio() };
}
host.addEventListener(
  'pointerdown',
  (event) => {
    if (office.getState().presenting || event.button !== 0 || busy || drag) return;
    if (
      event
        .composedPath()
        .some(
          (node) =>
            node instanceof HTMLElement &&
            (node.classList.contains('drawing-guide') ||
              node.classList.contains('animation-marker')),
        )
    )
      return;
    selectedAnimation = null;
    if (textEdit) return;
    if (
      !tool &&
      event.composedPath().some((node) => node instanceof HTMLElement && node.dataset.shapeId)
    )
      return;
    event.preventDefault();
    const initialSelection = [...selected];
    if (!tool && !event.shiftKey && !event.metaKey) selected.clear();
    render();
    focusCanvas();
    const snapped = lineTool(tool) ? connectionAt(point(event), undefined, event.altKey) : null;
    const p = snapped?.point ?? point(event);
    const box = document.createElement('div');
    box.className = 'draw-box';
    layer.append(box);
    drawing = {
      ...p,
      pointer: event.pointerId,
      box,
      preset: tool,
      selection: initialSelection,
      additive: event.shiftKey || event.metaKey,
      startConnection: snapped?.connection ?? null,
    };
    host.setPointerCapture(event.pointerId);
  },
  true,
);
host.addEventListener('pointermove', (event) => {
  if (!drawing || drawing.pointer !== event.pointerId) {
    if (lineTool(tool) && !drag) connectionAt(point(event), undefined, event.altKey);
    return;
  }
  const p = drawingPoint(event);
  const bounds = {
    x: Math.min(p.x, drawing.x),
    y: Math.min(p.y, drawing.y),
    w: Math.abs(p.x - drawing.x),
    h: Math.abs(p.y - drawing.y),
  };
  setBox(drawing.box, bounds);
  if (lineTool(drawing.preset)) {
    drawing.box.style.border = 'none';
    drawing.box.style.background = 'transparent';
    const w = bounds.w * ratio(),
      h = bounds.h * ratio();
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.style.cssText = `position:absolute;overflow:visible;inset:0;width:${Math.max(1, w)}px;height:${Math.max(1, h)}px`;
    const path = document.createElementNS(svg.namespaceURI, 'path');
    path.setAttribute('d', linePath(drawing.preset!, w, h));
    path.setAttribute('fill', 'none');
    path.setAttribute('stroke', '#4472C4');
    path.setAttribute('stroke-width', '1.5');
    path.setAttribute(
      'transform',
      `translate(${p.x < drawing.x ? w : 0} ${p.y < drawing.y ? h : 0}) scale(${p.x < drawing.x ? -1 : 1} ${p.y < drawing.y ? -1 : 1})`,
    );
    svg.append(path);
    drawing.box.replaceChildren(svg);
  }
});
host.addEventListener('pointerup', (event) => {
  if (!drawing || drawing.pointer !== event.pointerId) return;
  const start = drawing;
  const p = drawingPoint(event);
  drawing = null;
  start.box.remove();
  const bounds = {
    x: Math.min(p.x, start.x),
    y: Math.min(p.y, start.y),
    w: Math.abs(p.x - start.x),
    h: Math.abs(p.y - start.y),
  };
  if (!start.preset) {
    selected.clear();
    if (start.additive) for (const id of start.selection) selected.add(id);
    if (bounds.w * ratio() >= 3 && bounds.h * ratio() >= 3) {
      for (const shape of slideModel()?.shapes ?? []) {
        if (!shape.bounds) continue;
        const b = shape.bounds;
        const angle = (shape.rotation * Math.PI) / 180;
        const w = Math.abs(b.w * Math.cos(angle)) + Math.abs(b.h * Math.sin(angle));
        const h = Math.abs(b.w * Math.sin(angle)) + Math.abs(b.h * Math.cos(angle));
        const x = b.x + (b.w - w) / 2,
          y = b.y + (b.h - h) / 2;
        if (
          x >= bounds.x &&
          y >= bounds.y &&
          x + w <= bounds.x + bounds.w &&
          y + h <= bounds.y + bounds.h
        )
          selected.add(shape.id);
      }
    }
    render();
    return;
  }
  if (lineTool(start.preset)) {
    if (Math.hypot(p.x - start.x, p.y - start.y) * ratio() < 5) {
      p.x = start.x + 2743200;
      p.y = start.y;
      start.endConnection = null;
    }
    tool = null;
    void perform('insert', {
      preset: start.preset,
      from: { x: start.x, y: start.y },
      toPoint: p,
      connections: { start: start.startConnection ?? null, end: start.endConnection ?? null },
    });
    return;
  }
  if (bounds.w * ratio() < 5) bounds.w = 2743200;
  if (bounds.h * ratio() < 5) bounds.h = 914400;
  const preset = start.preset;
  tool = null;
  void perform('insert', { preset, changes: { bounds } });
});
function cancelDrawing() {
  if (!drawing) return;
  selected.clear();
  for (const id of drawing.selection) selected.add(id);
  drawing?.box.remove();
  drawing = null;
  tool = null;
  render();
}
host.addEventListener('pointercancel', cancelDrawing);
host.addEventListener('lostpointercapture', cancelDrawing);
let clearPictureMenuPreview: (() => void) | undefined;
function pictureMenuPreview(changes: NonNullable<EditCommand['changes']>) {
  const state = office.getState();
  const ids = [...selected];
  clearPictureMenuPreview?.();
  clearPictureMenuPreview = previewPictureMenu({
    command: { type: 'update', slide: state.index, ids, changes },
    revision: state.revision,
    shadow,
    layer,
    current: () =>
      office.getState().revision === state.revision &&
      office.getState().index === state.index &&
      ids.length === selected.size &&
      ids.every((id) => selected.has(id)),
  });
  return clearPictureMenuPreview;
}
function menuAt(items: MenuItem[], element: Element) {
  const rect = element.getBoundingClientRect();
  office.menu(items, rect.left, rect.bottom, element);
}
function pictureSourceMenu(replacing: boolean): MenuItem[] {
  return [
    {
      label: 'Picture from File...',
      action: () => {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/png,image/jpeg,image/gif,image/webp';
        input.setAttribute('aria-label', replacing ? 'Change Picture' : 'Insert Picture');
        input.hidden = true;
        document.body.append(input);
        const slide = office.getState().index;
        const revision = office.getState().revision;
        const ids = [...selected];
        input.addEventListener('cancel', () => input.remove(), { once: true });
        input.addEventListener(
          'change',
          async () => {
            try {
              const file = input.files?.[0];
              if (!file) return;
              if (file.size > 20_000_000) throw new Error('Picture must be under 20 MB.');
              const data = await new Promise<string>((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = () => resolve(String(reader.result).split(',')[1]!);
                reader.onerror = () => reject(new Error('Could not read the picture.'));
                reader.readAsDataURL(file);
              });
              if (office.getState().index !== slide || office.getState().revision !== revision)
                throw new Error('The slide changed. Choose the picture again.');
              if (replacing) {
                if (ids.length !== selected.size || !ids.every((id) => selected.has(id)))
                  throw new Error('The selection changed. Choose the picture again.');
                await send({
                  type: 'picture-replace',
                  slide,
                  ids,
                  image: { base64: data, name: file.name },
                });
                return;
              }
              const size = model()!;
              const w = size.width * 0.6,
                h = size.height * 0.6;
              await perform('insert', {
                preset: 'picture',
                image: { base64: data, name: file.name },
                changes: {
                  bounds: { x: (size.width - w) / 2, y: (size.height - h) / 2, w, h },
                },
              });
            } catch (cause) {
              error(cause instanceof Error ? cause.message : String(cause));
            } finally {
              input.remove();
            }
          },
          { once: true },
        );
        input.click();
      },
    },
  ];
}
function tableDeleteMenu(): MenuItem[] {
  const table = selection()[0];
  if (!table?.table) return [];
  const range = cellSelection?.range;
  const cell = textEdit?.cell ?? (range ? { row: range.row, column: range.column } : undefined);
  return [
    ...(['columns', 'rows'] as const).map((axis) => ({
      label: axis === 'rows' ? 'Delete Rows' : 'Delete Columns',
      action: async () => {
        const slideKey = slideModel()?.key;
        if (
          !(await perform('table-delete', {
            tableAxis: axis,
            ...(range ? { tableRange: range } : {}),
            ...(cell ? { cell: { ...cell, edits: [] } } : {}),
          })) ||
          !cell ||
          slideModel()?.key !== slideKey
        )
          return;
        const current = selection().find((shape) => shape.id === table.id);
        if (!current?.table) return;
        const row = Math.min(cell.row, current.table.cells.length - 1);
        const column = Math.min(cell.column, current.table.columnWidths.length - 1);
        for (let r = 0; r <= row; r++)
          for (let c = 0; c <= column; c++) {
            const span = current.table.cells[r]![c]!.span;
            if (
              !span.hMerge &&
              !span.vMerge &&
              r + span.rowSpan > row &&
              c + span.gridSpan > column
            ) {
              startCell(current, r, c);
              return;
            }
          }
      },
    })),
    { label: 'Delete Table', action: () => void perform('delete') },
  ];
}

let arrangeToSlide = false;
const usesSlideAlignment = () => arrangeToSlide || selected.size < 2;
function arrangeMenu(nativeMenu = false): MenuItem[] {
  const disabled = busy || !selection().length;
  return [
    ...(['front', 'back', 'forward', 'backward'] as const).map((order, i) => ({
      label: ['Bring to Front', 'Send to Back', 'Bring Forward', 'Send Backward'][i]!,
      shortcut: ['⇧⌘F', '⇧⌘B', '⌥⇧⌘F', '⌥⇧⌘B'][i]!,
      disabled,
      action: () => void perform('order', { order }),
    })),
    null,
    {
      label: 'Group',
      shortcut: '⌥⌘G',
      disabled: busy || selected.size < 2,
      action: () => void perform('group'),
    },
    {
      label: 'Ungroup',
      shortcut: '⌥⇧⌘G',
      disabled: busy || !selection().length || selection().some((s) => s.kind !== 'group'),
      action: () => void perform('ungroup'),
    },
    {
      label: 'Regroup',
      shortcut: '⌥⌘J',
      disabled: busy || !selection().some((shape) => shape.regroupIds?.length),
      action: () => void perform('regroup'),
    },
    null,
    {
      label: 'Rotate or Flip',
      disabled,
      children: [
        {
          label: 'Rotate Left 90°',
          action: () => void perform('update', { changes: { rotationDelta: -90 } }),
        },
        {
          label: 'Rotate Right 90°',
          action: () => void perform('update', { changes: { rotationDelta: 90 } }),
        },
        {
          label: 'Flip Horizontal',
          action: () => void perform('update', { changes: { flipToggle: 'horizontal' } }),
        },
        {
          label: 'Flip Vertical',
          action: () => void perform('update', { changes: { flipToggle: 'vertical' } }),
        },
        null,
        { label: 'More Rotation Options...', action: () => shapeProperties.show() },
      ],
    },
    {
      label: nativeMenu ? 'Align or Distribute' : 'Align',
      disabled,
      children: [
        ...['Left', 'Center', 'Right', 'Top', 'Middle', 'Bottom'].map((label, i) => ({
          label: 'Align ' + label,
          action: () => alignShapes(i),
        })),
        null,
        {
          label: 'Distribute Horizontally',
          disabled: selected.size < (usesSlideAlignment() ? 2 : 3),
          action: () => distributeShapes('x'),
        },
        {
          label: 'Distribute Vertically',
          disabled: selected.size < (usesSlideAlignment() ? 2 : 3),
          action: () => distributeShapes('y'),
        },
        null,
        {
          label: 'Align to Slide',
          checked: usesSlideAlignment(),
          action: () => {
            arrangeToSlide = true;
          },
        },
        {
          label: 'Align Selected Objects',
          checked: !usesSlideAlignment(),
          disabled: selected.size < 2,
          action: () => {
            arrangeToSlide = false;
          },
        },
      ],
    },
    null,
    {
      label: 'Selection Pane...',
      action: () => {
        byId('selection-pane').hidden = false;
        byId('animation-pane').hidden = true;
        renderSelection();
      },
    },
  ];
}
function distributeShapes(axis: 'x' | 'y') {
  const positions = distributedPositions(
    selection(),
    axis,
    usesSlideAlignment() ? model() : undefined,
  );
  if (positions.length) void perform('update', { positions });
}
function alignShapes(direction: number) {
  const size = model();
  if (!size) return;
  const positions = alignedPositions(selection(), direction, size, usesSlideAlignment());
  if (positions.length) void perform('update', { positions });
}
const searchField = byId<HTMLInputElement>('find-text');
let lastSearch: { query: string; match: TextMatch } | null = null;
let searching = false;
const searchOptions: TextSearchOptions = {};
function focusSearch() {
  searchField.focus();
  searchField.select();
}
async function searchText(backward = false, after?: TextMatch) {
  if (searching || busy || !searchField.value) return;
  searching = true;
  try {
    if (!(await finishText())) return;
    const query = searchField.value;
    const matches = findText(model()?.slides ?? [], query, searchOptions);
    if (!matches.length) {
      lastSearch = null;
      byId('find-status').textContent = 'No matches';
      searchField.focus();
      return;
    }
    const previous =
      lastSearch?.query === query
        ? matches.findIndex(
            (match) =>
              match.slide === lastSearch!.match.slide &&
              match.shape === lastSearch!.match.shape &&
              match.start === lastSearch!.match.start &&
              match.end === lastSearch!.match.end,
          )
        : -1;
    let index =
      previous >= 0
        ? (previous + (backward ? -1 : 1) + matches.length) % matches.length
        : backward
          ? matches.reduce(
              (found, match, position) =>
                match.slide <= office.getState().index ? position : found,
              -1,
            )
          : matches.findIndex((match) => match.slide >= office.getState().index);
    if (after) {
      const shapes = model()?.slides[after.slide]?.shapes ?? [];
      const shapeIndex = shapes.findIndex((shape) => shape.id === after.shape);
      index = matches.findIndex(
        (match) =>
          match.slide > after.slide ||
          (match.slide === after.slide &&
            (shapes.findIndex((shape) => shape.id === match.shape) > shapeIndex ||
              (match.shape === after.shape && match.start >= after.end))),
      );
    }
    if (index < 0) index = backward ? matches.length - 1 : 0;
    const match = matches[index]!;
    lastSearch = { query, match };
    office.selectSlide(match.slide);
    const shape = allShapes().find((shape) => shape.id === match.shape);
    if (shape) startText(shape, match);
    byId('find-status').textContent = index + 1 + ' of ' + matches.length;
  } finally {
    searching = false;
  }
}
async function openReplace() {
  const existing = document.querySelector<HTMLDialogElement>('.find-replace-dialog');
  if (existing) {
    existing.querySelector<HTMLInputElement>('input')!.focus();
    return;
  }
  if (!(await finishText())) return;
  const dialog = document.createElement('dialog');
  dialog.className = 'section-dialog find-replace-dialog';
  dialog.setAttribute('aria-label', 'Find and Replace');
  dialog.innerHTML = `<form><h2>Find and Replace</h2><label>Find what:<input name="find" autocomplete="off"></label><label>Replace with:<input name="replacement" autocomplete="off"></label><label class="search-option"><input type="checkbox" name="matchCase">Match case</label><label class="search-option"><input type="checkbox" name="wholeWords">Find whole words only</label><p role="status" aria-live="polite"></p><div><button type="button" data-close>Close</button><button type="button" data-replace-all>Replace All</button><button type="button" data-replace>Replace</button><button type="submit">Find Next</button></div></form>`;
  const find = dialog.querySelector<HTMLInputElement>('[name=find]')!;
  const replacement = dialog.querySelector<HTMLInputElement>('[name=replacement]')!;
  const status = dialog.querySelector<HTMLElement>('[role=status]')!;
  find.value = searchField.value;
  for (const option of ['matchCase', 'wholeWords'] as const) {
    const checkbox = dialog.querySelector<HTMLInputElement>('[name=' + option + ']')!;
    checkbox.checked = searchOptions[option] ?? false;
    checkbox.onchange = () => {
      searchOptions[option] = checkbox.checked;
      lastSearch = null;
      status.textContent = '';
      byId('find-status').textContent = '';
    };
  }
  let pending = false;
  const syncQuery = () => {
    if (searchField.value !== find.value) lastSearch = null;
    searchField.value = find.value;
  };
  const run = async (replace: boolean) => {
    if (pending || !find.value || busy) return;
    pending = true;
    for (const button of dialog.querySelectorAll<HTMLButtonElement | HTMLInputElement>(
      'button, input',
    ))
      button.disabled = true;
    try {
      syncQuery();
      const match = lastSearch?.query === find.value ? lastSearch.match : undefined;
      const edit = textEdit;
      if (
        replace &&
        match &&
        edit &&
        edit.slide === match.slide &&
        edit.shape.id === match.shape &&
        edit.field.selectionStart === match.start &&
        edit.field.selectionEnd === match.end &&
        findText(
          [{ shapes: [{ ...edit.shape, text: edit.value }] }],
          find.value,
          searchOptions,
        ).some((candidate) => candidate.start === match.start && candidate.end === match.end)
      ) {
        edit.field.replaceSelection(replacement.value);
        const after = { ...match, end: match.start + replacement.value.length };
        if (!(await finishText())) {
          status.textContent = 'Could not replace text. Please try again.';
          return;
        }
        await searchText(false, after);
        status.textContent = 'Replaced. ' + byId('find-status').textContent;
      } else {
        await searchText();
        status.textContent = byId('find-status').textContent;
      }
    } finally {
      pending = false;
      for (const button of dialog.querySelectorAll<HTMLButtonElement | HTMLInputElement>(
        'button, input',
      ))
        button.disabled = false;
    }
  };
  dialog.querySelector('form')!.onsubmit = (event) => {
    event.preventDefault();
    void run(false);
  };
  dialog.querySelector<HTMLButtonElement>('[data-replace]')!.onclick = () => void run(true);
  dialog.querySelector<HTMLButtonElement>('[data-replace-all]')!.onclick = async () => {
    if (pending || !find.value || busy) return;
    pending = true;
    for (const button of dialog.querySelectorAll<HTMLButtonElement | HTMLInputElement>(
      'button, input',
    ))
      button.disabled = true;
    try {
      if (!(await finishText())) return;
      syncQuery();
      const query = find.value;
      const count = findText(model()?.slides ?? [], query, searchOptions).length;
      if (!count) {
        status.textContent = 'No matches';
        return;
      }
      const saved = await send({
        type: 'text-replace-all',
        slide: office.getState().index,
        replacement: { query, text: replacement.value, options: { ...searchOptions } },
      });
      if (saved) {
        lastSearch = null;
        byId('find-status').textContent = '';
        status.textContent = 'Replaced ' + count + (count === 1 ? ' occurrence.' : ' occurrences.');
      } else status.textContent = 'Could not replace text. Please try again.';
    } finally {
      pending = false;
      for (const button of dialog.querySelectorAll<HTMLButtonElement | HTMLInputElement>(
        'button, input',
      ))
        button.disabled = false;
    }
  };
  const close = () => {
    if (!pending) {
      dialog.close();
      dialog.remove();
      if (textEdit) textEdit.field.focus();
      else focusCanvas();
    }
  };
  dialog.querySelector<HTMLButtonElement>('[data-close]')!.onclick = close;
  dialog.onkeydown = (event) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      close();
    }
  };
  document.body.append(dialog);
  dialog.show();
  find.focus();
  find.select();
}
searchField.addEventListener('input', () => {
  lastSearch = null;
  byId('find-status').textContent = '';
});
searchField.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') {
    event.preventDefault();
    void searchText(event.shiftKey);
  }
  if (event.key === 'Escape') {
    event.preventDefault();
    byId('find-status').textContent = '';
    focusCanvas();
  }
});
byId('find-menu').onclick = () =>
  menuAt([{ label: 'Find and Replace', action: () => void openReplace() }], byId('find-menu'));
byId('find-next').onclick = () => void searchText();
byId('find-previous').onclick = () => void searchText(true);
document.addEventListener(
  'keydown',
  (event) => {
    if (
      office.getState().presenting ||
      document.querySelector('dialog[open]') ||
      event.altKey ||
      !(event.metaKey || event.ctrlKey)
    )
      return;
    const key = event.key.toLowerCase();
    if (key === 'k' && !event.shiftKey && !cellSelection && selected.size === 1) {
      event.preventDefault();
      event.stopImmediatePropagation();
      void action('link', host);
    }
    if (key === 'f' && !event.shiftKey) {
      event.preventDefault();
      event.stopImmediatePropagation();
      focusSearch();
    }
    if (key === 'g' && searchField.value) {
      event.preventDefault();
      event.stopImmediatePropagation();
      void searchText(event.shiftKey);
    }
  },
  true,
);
async function textClipboardAction(name: 'copy' | 'cut' | 'paste') {
  const edit = textEdit;
  if (!edit || clipboardPending) return;
  const field = edit.field;
  const start = field.selectionStart;
  const end = field.selectionEnd;
  const value = edit.value;
  if (name !== 'paste' && start === end) return;
  clipboardPending = true;
  updateControls();
  const unchanged = () =>
    textEdit === edit &&
    edit.value === value &&
    field.selectionStart === start &&
    field.selectionEnd === end;
  try {
    error('');
    if (name === 'paste') {
      const text = await navigator.clipboard.readText();
      if (!text) return;
      if (!unchanged()) throw new Error('The text selection changed. Paste again.');
      field.replaceSelection(text);
    } else {
      await navigator.clipboard.writeText(value.slice(start, end));
      if (name === 'cut') {
        if (!unchanged())
          throw new Error('The text selection changed. Text was copied but not cut.');
        field.deleteSelection();
      } else if (unchanged()) {
        field.focus();
        field.setSelectionRange(start, end);
      }
    }
  } catch (cause) {
    error(cause instanceof Error ? cause.message : 'Could not access the clipboard.');
  } finally {
    clipboardPending = false;
    updateControls();
  }
}
function textFormatShortcut(event: KeyboardEvent, origin: Element): boolean {
  if (event.isComposing || !event.metaKey) return false;
  if (
    event.shiftKey &&
    !event.ctrlKey &&
    !event.altKey &&
    (event.code === 'Digit1' || event.key === '!')
  ) {
    if (busy || !selection().length) return false;
    event.preventDefault();
    event.stopImmediatePropagation();
    shapeProperties.show();
    return true;
  }
  const key = event.key.toLowerCase();
  let name: string | undefined;
  if (event.ctrlKey && !event.altKey) {
    if (event.code === 'Equal' || key === '=' || key === '+')
      name = event.shiftKey ? 'superscript' : 'subscript';
  } else if (!event.ctrlKey && !event.shiftKey) {
    if (event.altKey) {
      if (event.code === 'KeyM' || key === 'm') name = 'paragraph-dialog';
    } else {
      name = {
        t: 'font-dialog',
        e: 'align-center',
        j: 'align-justify',
        l: 'align-left',
        r: 'align-right',
        '[': 'promote',
        ']': 'demote',
      }[key];
    }
  }
  if (!name || !activeShape()?.textable || busy) return false;
  event.preventDefault();
  event.stopImmediatePropagation();
  void action(name, origin);
  return true;
}

function textLinkAt(shape: EditorShape, range: { start: number; end: number }) {
  const merged: typeof shape.textLinks = [];
  for (const item of shape.textLinks) {
    const previous = merged.at(-1);
    if (
      previous &&
      previous.end <= item.start &&
      /^\n*$/.test(shape.text.slice(previous.end, item.start)) &&
      JSON.stringify(previous.link) === JSON.stringify(item.link)
    )
      previous.end = item.end;
    else merged.push({ ...item });
  }
  return merged.find((item) =>
    range.start === range.end
      ? item.start <= range.start && range.start < item.end
      : item.start <= range.start && item.end >= range.end,
  );
}

function selectedHyperlink() {
  if (selected.size !== 1 || cellSelection) return null;
  if (!textEdit) return activeShape()?.link ?? null;
  return (
    textLinkAt(textEdit.shape, {
      start: textEdit.field.selectionStart,
      end: textEdit.field.selectionEnd,
    })?.link ?? null
  );
}

async function openSelectedHyperlink() {
  if (busy) return;
  const link = selectedHyperlink();
  if (!link) return;
  const action = link.action;
  if (action.kind === 'url') {
    try {
      const url = new URL(action.url, location.href);
      if (!['http:', 'https:', 'mailto:', 'tel:'].includes(url.protocol)) {
        error('This link cannot be opened in the browser.');
        return;
      }
      window.open(url.href, '_blank', 'noopener,noreferrer');
    } catch {
      error('This link has an invalid address.');
    }
    return;
  }
  if (action.kind === 'endShow' || action.kind === 'lastSlideViewed') return;
  const key = slideModel()?.key;
  if (!(await finishText()) || slideModel()?.key !== key) return;
  if (action.kind === 'customShow') {
    office.startCustomShow(action.id);
    return;
  }
  const slides = model()?.slides ?? [];
  const index = office.getState().index;
  const destination =
    action.kind === 'slide'
      ? slides.findIndex((slide) => slide.key === action.slide)
      : action.kind === 'firstSlide'
        ? 0
        : action.kind === 'lastSlide'
          ? slides.length - 1
          : action.kind === 'nextSlide'
            ? index + 1
            : index - 1;
  if (destination >= 0 && destination < slides.length) office.selectSlide(destination, true);
}

function openHyperlinkMenu(): MenuItem[] {
  return selectedHyperlink()
    ? [{ label: 'Open Hyperlink', disabled: busy, action: () => void openSelectedHyperlink() }]
    : [];
}

async function chooseTransition(value: EffectChoice) {
  const key = slideModel()?.key;
  if (
    busy ||
    pendingTransitionDuration ||
    pendingAdvanceTiming ||
    !(await finishText()) ||
    slideModel()?.key !== key
  )
    return;
  if (await perform('transition-effect', { transitionEffect: value })) office.previewTransition();
}

async function action(name: string, element: Element) {
  if (name.startsWith('animation-preview')) {
    if (name === 'animation-preview-stop') office.stopAnimationPreview();
    else if (!busy && !pendingAnimationTiming) {
      const slideKey = slideModel()?.key;
      const animationIds = selectedAnimationIds();
      if ((await finishText()) && slideKey === slideModel()?.key)
        office.previewAnimations(
          name === 'animation-preview' ? undefined : animationIds,
          name === 'animation-preview-selected',
        );
    }
    return;
  }
  if (name === 'animation-earlier' || name === 'animation-later') {
    const direction = name === 'animation-earlier' ? 'earlier' : 'later';
    if (!canMoveAnimation(direction) || !selectedAnimation) return;
    const target = selectedAnimation;
    const animationIds = selectedAnimationIds();
    if (!(await finishText()) || slideModel()?.key !== target.slide) return;
    if (await perform('animation-move', { animationIds, animationDirection: direction })) {
      selectedAnimation = target;
      updateControls();
      [...byId('animation-list').querySelectorAll<HTMLButtonElement>('button')]
        .find((button) => button.dataset.animationId === target.id)
        ?.focus();
    }
    return;
  }
  if (name === 'animation-remove') {
    if (busy || pendingAnimationTiming || !selectedAnimation) return;
    const target = selectedAnimation;
    const animationIds = selectedAnimationIds();
    const effects = slideModel()?.animations.flat() ?? [];
    const position = effects.findIndex((effect) => animationIds.includes(effect.timingId ?? ''));
    if (!(await finishText()) || slideModel()?.key !== target.slide) return;
    if (await perform('animation-remove', { animationIds })) {
      const remaining = slideModel()?.animations.flat() ?? [];
      const next = remaining[Math.min(position, remaining.length - 1)];
      selectedAnimation = next?.timingId ? { slide: target.slide, id: next.timingId } : null;
      updateControls();
      if (element.classList.contains('animation-marker')) {
        const marker = [...layer.querySelectorAll<HTMLButtonElement>('.animation-marker')].find(
          (button) => button.dataset.animationId === next?.timingId,
        );
        if (marker) marker.focus();
        else focusCanvas();
        return;
      }
      const buttons = [...byId('animation-list').querySelectorAll<HTMLButtonElement>('button')];
      (
        buttons.find((button) => button.dataset.animationId === next?.timingId) ??
        byId('animation-pane-close')
      ).focus();
    }
    return;
  }
  if (name === 'animation-pane') {
    const pane = byId('animation-pane');
    pane.hidden = !pane.hidden;
    if (!pane.hidden) byId('selection-pane').hidden = true;
    updateControls();
    return;
  }
  if (name === 'animation-add') {
    if (busy || pendingAnimationTiming || (!selected.size && !activeAnimation())) return;
    const animationIds = selectedAnimationIds();
    const target = animationIds.length > 0;
    const slideKey = slideModel()?.key;
    if (!(await finishText()) || slideModel()?.key !== slideKey) return;
    const effect = (element as HTMLElement).dataset.effect;
    if (
      effect === 'appear' ||
      effect === 'disappear' ||
      effect === 'fadeIn' ||
      effect === 'fadeOut'
    )
      await perform(
        target ? 'animation-settings' : 'animation-add',
        target ? { animationIds, animationSettings: { effect } } : { animation: { effect } },
      );
    return;
  }
  if (name === 'transition-preview') {
    if (busy || pendingTransitionDuration || pendingAdvanceTiming || !(await finishText())) return;
    office.previewTransition();
    return;
  }
  if (name === 'transition-effect') {
    const effect = (element as HTMLElement).dataset.effect;
    if (effect) await chooseTransition({ effect });
    return;
  }
  if (name === 'transition-more') {
    menuAt(
      transitionEffects.map(([effect, label]) => ({
        label,
        checked: (slideModel()?.transition?.effect ?? 'none') === effect,
        action: () => void chooseTransition({ effect }),
      })),
      element,
    );
    return;
  }
  if (name === 'transition-options') {
    const current = slideModel()?.transition;
    menuAt(
      transitionOptions(current?.effect ?? 'none').map(({ label, value }) => ({
        label,
        checked: transitionOptionSelected(current, value),
        action: () => void chooseTransition(value),
      })),
      element,
    );
    return;
  }
  if (name === 'transition-apply-all') {
    const key = slideModel()?.key;
    if (
      busy ||
      pendingTransitionDuration ||
      pendingAdvanceTiming ||
      !(await finishText()) ||
      slideModel()?.key !== key
    )
      return;
    await perform('transition-apply-all');
    return;
  }
  if (name === 'show-properties') {
    if (busy || !(await finishText())) return;
    const editor = model();
    if (!editor) return;
    const revision = office.getState().revision;
    openShowProperties(
      editor.showProperties,
      editor.slides.length,
      editor.customShows,
      async (showProperties) => {
        if (office.getState().revision !== revision)
          throw new Error('The presentation changed. Open Set Up Show again.');
        if (JSON.stringify(showProperties) === JSON.stringify(model()?.showProperties)) return true;
        return perform('show-properties', { showProperties });
      },
      focusCanvas,
    );
    return;
  }
  if (name === 'custom-shows') {
    if (busy || !(await finishText())) return;
    const editor = model();
    if (!editor) return;
    const rect = element.getBoundingClientRect();
    office.menu(
      [
        ...editor.customShows.map((show) => ({
          label: show.name,
          disabled: !show.slides.length,
          action: () => office.startCustomShow(show.id),
        })),
        ...(editor.customShows.length ? [null] : []),
        {
          label: 'Custom Slide Show...',
          action: () => {
            let revision = office.getState().revision;
            const current = model();
            if (!current) return;
            openCustomShows(
              current.customShows,
              current.slides.map((slide) => ({
                key: slide.key,
                title:
                  slide.shapes.find(
                    (shape) => shape.placeholder === 'title' || shape.placeholder === 'ctrTitle',
                  )?.text ?? '',
              })),
              async (shows) => {
                if (office.getState().revision !== revision)
                  throw new Error('The presentation changed. Open Custom Shows again.');
                if (JSON.stringify(shows) === JSON.stringify(model()?.customShows)) return true;
                const saved = await perform('custom-shows', { customShows: shows });
                if (saved) revision = office.getState().revision;
                return saved;
              },
              office.startCustomShow,
              focusCanvas,
            );
          },
        },
      ],
      rect.left,
      rect.bottom,
      element,
    );
    return;
  }
  if (name === 'action-settings') {
    if (busy || cellSelection || selected.size !== 1) return;
    const slideKey = slideModel()?.key;
    if (!(await finishText()) || slideModel()?.key !== slideKey) return;
    const shape = selection()[0];
    if (!shape) return;
    const revision = office.getState().revision;
    openActionDialog(
      { click: shape.link, hover: shape.hoverLink, sounds: shape.actionSounds },
      (model()?.slides ?? []).map((slide) => ({
        key: slide.key,
        title:
          slide.shapes.find(
            (item) => item.placeholder === 'title' || item.placeholder === 'ctrTitle',
          )?.text ?? '',
      })),
      async (links) => {
        if (
          slideModel()?.key !== slideKey ||
          office.getState().revision !== revision ||
          selected.size !== 1 ||
          !selected.has(shape.id)
        ) {
          throw new Error('The selection changed. Open Action Settings again.');
        }
        if (
          JSON.stringify(links) ===
          JSON.stringify({ click: shape.link, hover: shape.hoverLink, sounds: shape.actionSounds })
        )
          return true;
        return perform('object-actions', {
          link: links.click,
          hoverLink: links.hover,
          actionSounds: links.sounds,
        });
      },
      focusCanvas,
      model()?.customShows ?? [],
    );
    return;
  }

  if (name === 'link') {
    if (busy || cellSelection || selected.size !== 1) return;
    const editingCell = textEdit?.cell;
    const selectedShape = activeShape();
    if (!selectedShape) return;
    const slideKey = slideModel()?.key;
    const range = textEdit && {
      start: textEdit.field.selectionStart,
      end: textEdit.field.selectionEnd,
    };
    let restoredRange = range && { ...range };
    if (!(await finishText()) || slideModel()?.key !== slideKey || !selected.has(selectedShape.id))
      return;
    const currentShape = activeShape();
    const shape =
      currentShape && editingCell
        ? tableCellShape(currentShape, editingCell.row, editingCell.column)
        : currentShape;
    if (!shape) return;
    const revision = office.getState().revision;
    let initial = shape.link;
    if (range) {
      const existing = textLinkAt(shape, range);
      initial = existing?.link ?? null;
      if (existing && range.start === range.end) {
        range.start = existing.start;
        range.end = existing.end;
      }
    }
    const restore = () => {
      if (range && slideModel()?.key === slideKey && selected.has(shape.id)) {
        const current = allShapes().find((item) => item.id === shape.id);
        if (current) {
          if (editingCell)
            startCell(current, editingCell.row, editingCell.column, restoredRange ?? range);
          else startText(current, restoredRange ?? range);
        }
      } else focusCanvas();
    };
    openLinkDialog(
      initial,
      (model()?.slides ?? []).map((slide) => ({
        key: slide.key,
        title:
          slide.shapes.find(
            (item) => item.placeholder === 'title' || item.placeholder === 'ctrTitle',
          )?.text ?? '',
      })),
      async (link, displayText) => {
        if (
          slideModel()?.key !== slideKey ||
          office.getState().revision !== revision ||
          selected.size !== 1 ||
          !selected.has(shape.id)
        ) {
          error('The selection changed. Open Hyperlink again.');
          return false;
        }
        const ok = await perform(range ? 'text-link' : 'object-link', {
          link,
          ...(editingCell ? { cell: { ...editingCell, edits: [] } } : {}),
          ...(range
            ? { range: { ...range }, ...(displayText !== undefined ? { text: displayText } : {}) }
            : {}),
        });
        if (ok && range)
          restoredRange = {
            start: range.start,
            end: displayText !== undefined ? range.start + displayText.length : range.end,
          };
        return ok;
      },
      restore,
      range ? shape.text.slice(range.start, range.end) : undefined,
      model()?.customShows ?? [],
    );
    return;
  }

  if (name === 'table-select') {
    menuAt(tableSelectMenu(), element);
    return;
  }
  if (name === 'view-menu') {
    const sorter = document.body.classList.contains('slide-sorter');
    const disabled = !model()?.slides.length;
    menuAt(
      [
        { label: 'Normal', shortcut: '⌘1', checked: !sorter, action: () => setDocumentView(false) },
        {
          label: 'Slide Sorter',
          shortcut: '⌘2',
          checked: sorter,
          action: () => setDocumentView(true),
        },
        {
          label: 'Presenter View',
          shortcut: '⌥↩',
          disabled,
          action: () => byId('presenter-view-start').click(),
        },
        { label: 'Slide Show', shortcut: '⇧⌘↩', disabled, action: () => byId('present').click() },
        null,
        {
          label: 'Ribbon',
          shortcut: '⌥⌘R',
          checked: !document.body.classList.contains('ribbon-collapsed'),
          action: () => byId('collapse-ribbon').click(),
        },
        null,
        {
          label: 'Grid and Guides',
          children: gridMenu(),
        },
        {
          label: 'Zoom',
          disabled,
          children: [
            { label: 'Fit to Window', action: () => byId('view-fit').click() },
            { label: 'Zoom In', action: () => byId('zoom-in').click() },
            { label: 'Zoom Out', action: () => byId('zoom-out').click() },
            null,
            { label: 'Zoom...', action: () => byId('view-zoom').click() },
          ],
        },
      ],
      element,
    );
    return;
  }
  if (name === 'format-menu') {
    const disabled = busy || !activeShape()?.textable;
    menuAt(
      [
        {
          label: 'Font...',
          shortcut: '⌘T',
          disabled,
          action: () => void action('font-dialog', element),
        },
        {
          label: 'Paragraph...',
          shortcut: '⌥⌘M',
          disabled,
          action: () => void action('paragraph-dialog', element),
        },
      ],
      element,
    );
    return;
  }
  if (name === 'edit-menu') {
    const state = office.getState();
    menuAt(
      [
        {
          label: 'Undo',
          shortcut: '⌘Z',
          disabled: !(textEdit?.undoCount || state.history?.undo),
          action: () => void action('undo', element),
        },
        {
          label: 'Redo',
          shortcut: '⌘Y',
          disabled: !(textEdit?.redoCount || state.history?.redo),
          action: () => void action('redo', element),
        },
        null,
        ...(['cut', 'copy', 'paste'] as const).map((name) => ({
          label: { cut: 'Cut', copy: 'Copy', paste: 'Paste' }[name],
          shortcut: { cut: '⌘X', copy: '⌘C', paste: '⌘V' }[name],
          disabled:
            clipboardPending ||
            (name === 'paste'
              ? textEdit
                ? !navigator.clipboard?.readText
                : !clipboard
              : textEdit
                ? textEdit.field.selectionStart === textEdit.field.selectionEnd
                : !selected.size),
          action: () => void action(name, element),
        })),
        null,
        {
          label: 'Duplicate',
          shortcut: '⌘D',
          disabled: !!textEdit || !selected.size,
          action: () => void perform('duplicate'),
        },
        {
          label: 'Delete',
          disabled: textEdit
            ? textEdit.field.selectionStart === textEdit.field.selectionEnd
            : !selected.size,
          action: () => {
            if (textEdit) textEdit.field.deleteSelection();
            else void perform('delete');
          },
        },
        null,
        {
          label: 'Select All',
          shortcut: '⌘A',
          disabled: !slideModel()?.shapes.length,
          action: () => void action('select-all', element),
        },
        null,
        {
          label: 'Find',
          children: [
            { label: 'Find…', shortcut: '⌘F', action: focusSearch },
            { label: 'Replace…', action: () => void openReplace() },
            {
              label: 'Find Next',
              shortcut: '⌘G',
              disabled: !searchField.value,
              action: () => void searchText(),
            },
            {
              label: 'Find Previous',
              shortcut: '⇧⌘G',
              disabled: !searchField.value,
              action: () => void searchText(true),
            },
          ],
        },
      ],
      element,
    );
    return;
  }
  if ((name === 'copy' || name === 'cut' || name === 'paste') && textEdit) {
    await textClipboardAction(name);
    return;
  }
  if (name === 'select-all' && textEdit) {
    textEdit.field.focus();
    textEdit.field.select();
    return;
  }

  if ((name === 'undo' || name === 'redo') && textEdit) {
    if (textEdit.undoText(name === 'redo')) return;
    const editing = textEdit;
    const slideKey = slideModel()?.key;
    const range = { start: editing.field.selectionStart, end: editing.field.selectionEnd };
    if (!(await finishText())) return;
    if (!(await send(undefined, name)) || slideModel()?.key !== slideKey) return;
    const shape = allShapes().find((item) => item.id === editing.shape.id);
    if (shape) {
      if (editing.cell) startCell(shape, editing.cell.row, editing.cell.column, range);
      else startText(shape, range);
    }
    return;
  }
  if (name === 'paragraph-dialog') {
    const shape = activeShape();
    if (!shape || busy) return;
    const slideKey = slideModel()?.key;
    const revision = office.getState().revision;
    const range = textEdit && {
      start: textEdit.field.selectionStart,
      end: textEdit.field.selectionEnd,
    };
    const restore = () => {
      if (textEdit && range && textEdit.shape.id === shape.id && slideModel()?.key === slideKey) {
        textEdit.field.focus();
        textEdit.field.setSelectionRange(range.start, range.end);
      } else focusCanvas();
    };
    let paragraphs = textEdit
      ? textEdit.paragraphs.filter(
          (paragraph) =>
            range &&
            (range.start === range.end
              ? range.start >= paragraph.start && range.start <= paragraph.end
              : range.start <= paragraph.end && range.end > paragraph.start),
        )
      : selection().flatMap((item) => item.paragraphs);
    if (!textEdit && cellSelection) {
      const table = selection().find((item) => item.id === cellSelection!.id);
      const cells = cellSelection.range;
      paragraphs = [];
      if (table)
        for (let row = cells.row; row < cells.row + cells.rows; row++) {
          for (let column = cells.column; column < cells.column + cells.columns; column++) {
            paragraphs.push(...(tableCellShape(table, row, column)?.paragraphs ?? []));
          }
        }
    }
    const initial = paragraphs.map((paragraph) => paragraph.properties);
    openParagraphDialog(
      initial,
      async (paragraph) => {
        if (slideModel()?.key !== slideKey || office.getState().revision !== revision) {
          error('The slide changed. Open Paragraph again.');
          return false;
        }
        restore();
        return perform('update', { changes: { paragraph } });
      },
      restore,
    );
    return;
  }
  if (name === 'font-dialog' || name === 'font-spacing-dialog') {
    const shape = activeShape();
    if (!shape || busy) return;
    const slideKey = slideModel()?.key;
    const revision = office.getState().revision;
    const range = textEdit && {
      start: textEdit.field.selectionStart,
      end: textEdit.field.selectionEnd,
    };
    const restore = () => {
      if (textEdit && range && textEdit.shape.id === shape.id && slideModel()?.key === slideKey) {
        textEdit.field.focus();
        textEdit.field.setSelectionRange(range.start, range.end);
      } else focusCanvas();
    };
    let formats = [shape.format];
    if (textEdit && range && range.start !== range.end) {
      formats = textEdit.runs
        .filter((run) => run.start < range.end && run.end > range.start)
        .map((run) => run.format);
    } else if (!textEdit) {
      let targets = selection();
      if (cellSelection) {
        const table = targets.find((item) => item.id === cellSelection!.id);
        const cells = cellSelection.range;
        targets = [];
        if (table)
          for (let row = cells.row; row < cells.row + cells.rows; row++) {
            for (let column = cells.column; column < cells.column + cells.columns; column++) {
              const cell = tableCellShape(table, row, column);
              if (cell) targets.push(cell);
            }
          }
      }
      formats = targets.flatMap((item) =>
        item.runs.length ? item.runs.map((run) => run.format) : [item.format],
      );
    }
    openFontDialog(
      formats,
      async (format) => {
        if (slideModel()?.key !== slideKey || office.getState().revision !== revision) {
          error('The slide changed. Open Font again.');
          return false;
        }
        restore();
        return perform('update', { changes: { format } });
      },
      restore,
      name === 'font-spacing-dialog' ? 'spacing' : 'font',
    );
    return;
  }
  const characterAction = [
    'bold',
    'italic',
    'underline',
    'strike',
    'superscript',
    'subscript',
    'font-grow',
    'font-shrink',
    'character-spacing',
    'highlight',
    'change-case',
  ].includes(name);
  const paragraphAction =
    name.startsWith('align-') ||
    ['bullets', 'numbering', 'promote', 'demote', 'line-spacing'].includes(name);
  if (
    !characterAction &&
    !paragraphAction &&
    ![
      'text-anchor',
      'text-direction',
      'table-shading',
      'table-margins',
      'table-split',
      'table-borders',
      'table-delete',
      'table-insert-above',
      'table-insert-below',
      'table-insert-left',
      'table-insert-right',
      'table-distribute-rows',
      'table-distribute-columns',
    ].includes(name) &&
    name !== 'undo' &&
    name !== 'redo' &&
    !(await finishText())
  )
    return;
  const shape = activeShape();
  if (name === 'table-split') {
    const table = selection()[0];
    const range = cellSelection?.range;
    const target = textEdit?.cell ?? (range && { row: range.row, column: range.column });
    if (!table?.table || !target) return;
    const slide = office.getState().index;
    const dialog = document.createElement('dialog');
    dialog.className = 'section-dialog table-margins-dialog table-split-dialog';
    dialog.setAttribute('aria-label', 'Split Cells');
    dialog.innerHTML = `<form><h2>Split Cells</h2><label>Number of columns:<input name="columns" type="number" min="1" max="100" step="1" value="2" required></label><label>Number of rows:<input name="rows" type="number" min="1" max="100" step="1" value="1" required></label><div><button type="button" data-cancel>Cancel</button><button type="submit">Insert</button></div></form>`;
    const close = () => {
      dialog.close();
      dialog.remove();
      textEdit?.field.focus();
    };
    dialog.oncancel = (event) => {
      event.preventDefault();
      close();
    };
    dialog.querySelector<HTMLButtonElement>('[data-cancel]')!.onclick = close;
    dialog.querySelector('form')!.onsubmit = async (event) => {
      event.preventDefault();
      const columns = Number(dialog.querySelector<HTMLInputElement>('[name=columns]')!.value);
      const rows = Number(dialog.querySelector<HTMLInputElement>('[name=rows]')!.value);
      close();
      if (
        await perform('table-split', {
          cell: { ...target, edits: [] },
          tableSplit: { rows, columns },
        })
      ) {
        const updated = allShapes().find((s) => s.id === table.id);
        if (updated && office.getState().index === slide)
          startCell(updated, target.row, target.column);
      }
    };
    document.body.append(dialog);
    dialog.showModal();
    dialog.querySelector<HTMLInputElement>('[name=columns]')!.select();
    return;
  }
  if (name === 'table-merge') {
    const selection = cellSelection;
    if (!selection) return;
    if (await perform('table-merge', { tableRange: selection.range })) {
      const table = allShapes().find((s) => s.id === selection.id);
      if (table && office.getState().index === selection.slide)
        startCell(table, selection.range.row, selection.range.column);
    }
    return;
  }
  if (name === 'table-delete') {
    menuAt(tableDeleteMenu(), element);
    return;
  }
  if (name.startsWith('table-insert-')) {
    const tableShape = selection()[0];
    if (!tableShape?.table) return;
    const range = cellSelection?.range;
    const cell = textEdit?.cell ?? (range && { row: range.row, column: range.column });
    const placement = name.slice('table-insert-'.length) as NonNullable<
      EditCommand['tablePlacement']
    >;
    const rows = placement === 'above' || placement === 'below';
    const after = placement === 'below' || placement === 'right';
    const span = cell ? tableShape.table.cells[cell.row]![cell.column]!.span : undefined;
    const index = range
      ? (rows ? range.row : range.column) + (after ? (rows ? range.rows : range.columns) : 0)
      : cell
        ? (rows ? cell.row : cell.column) + (after ? (rows ? span!.rowSpan : span!.gridSpan) : 0)
        : after
          ? rows
            ? tableShape.table.rowHeights.length
            : tableShape.table.columnWidths.length
          : 0;
    const slideKey = slideModel()?.key;
    if (
      await perform('table-insert', {
        tablePlacement: placement,
        ...(range ? { tableRange: range } : {}),
        ...(cell ? { cell: { ...cell, edits: [] } } : {}),
      })
    ) {
      const current = selection().find((shape) => shape.id === tableShape.id);
      if (current && slideModel()?.key === slideKey)
        startCell(current, rows ? index : (cell?.row ?? 0), rows ? (cell?.column ?? 0) : index);
    }
    return;
  }
  if (name === 'table-distribute-rows' || name === 'table-distribute-columns') {
    await perform('table-distribute', {
      tableAxis: name === 'table-distribute-rows' ? 'rows' : 'columns',
    });
    return;
  }
  if (name === 'table-margins') {
    const table = selection()[0]?.table;
    if (!table) return;
    const target = textEdit?.cell;
    const cells = target ? [table.cells[target.row]![target.column]!] : table.cells.flat();
    const sides = ['top', 'bottom', 'left', 'right'] as const;
    const dialog = document.createElement('dialog');
    dialog.className = 'section-dialog table-margins-dialog';
    dialog.setAttribute('aria-label', 'Cell Margins');
    dialog.innerHTML = `<form><h2>Cell Margins</h2>${sides.map((side) => `<label>${side[0]!.toUpperCase() + side.slice(1)}<input name="${side}" aria-label="${side[0]!.toUpperCase() + side.slice(1)} margin" type="number" min="0" max="56" step="0.001"> in</label>`).join('')}<div><button type="button" data-cancel>Cancel</button><button type="submit">OK</button></div></form>`;
    const initial = new Map<string, string>();
    for (const side of sides) {
      const input = dialog.querySelector<HTMLInputElement>(`[name=${side}]`)!;
      const values = cells.map(
        (cell) => cell.margins[side] ?? (side === 'left' || side === 'right' ? 91440 : 45720),
      );
      const mixed = values.some((value) => value !== values[0]);
      input.value = mixed ? '' : String(Number((values[0]! / 914400).toFixed(3)));
      input.placeholder = mixed ? 'Mixed' : '';
      initial.set(side, input.value);
    }
    const range = textEdit && {
      start: textEdit.field.selectionStart,
      end: textEdit.field.selectionEnd,
    };
    const close = () => {
      dialog.close();
      dialog.remove();
      if (textEdit && range) {
        textEdit.field.focus();
        textEdit.field.setSelectionRange(range.start, range.end);
      }
    };
    dialog.oncancel = (event) => {
      event.preventDefault();
      close();
    };
    dialog.querySelector<HTMLButtonElement>('[data-cancel]')!.onclick = close;
    dialog.querySelector('form')!.onsubmit = (event) => {
      event.preventDefault();
      const margins: NonNullable<EditCommand['tableMargins']> = {};
      for (const side of sides) {
        const input = dialog.querySelector<HTMLInputElement>(`[name=${side}]`)!;
        if (input.value && input.value !== initial.get(side))
          margins[side] = Math.round(Number(input.value) * 914400);
      }
      close();
      if (Object.keys(margins).length) void perform('table-margins', { tableMargins: margins });
    };
    document.body.append(dialog);
    dialog.showModal();
    return;
  }
  if (name === 'table-borders') {
    menuAt(
      (
        [
          ['Bottom Border', 'bottom'],
          ['Top Border', 'top'],
          ['Left Border', 'left'],
          ['Right Border', 'right'],
          ['No Border', 'none'],
          ['All Borders', 'all'],
          ['Outside Borders', 'outside'],
          ['Inside Borders', 'inside'],
          ['Inside Horizontal Border', 'horizontal'],
          ['Inside Vertical Border', 'vertical'],
          ['Diagonal Down Border', 'tlToBr'],
          ['Diagonal Up Border', 'blToTr'],
        ] as const
      ).map(([label, mode]) => ({
        label,
        action: () =>
          void perform('table-borders', {
            tableBorders: {
              mode,
              color: byId<HTMLInputElement>('table-pen-color').value,
              weight: Number(byId<HTMLSelectElement>('table-pen-weight').value),
              dash: byId<HTMLSelectElement>('table-pen-style').value as NonNullable<
                EditCommand['tableBorders']
              >['dash'],
            },
          }),
      })),
      element,
    );
    return;
  }
  if (name === 'table-shading') {
    menuAt(
      [
        ...(
          [
            ['White', 'FFFFFF'],
            ['Black', '000000'],
            ['Dark Red', 'C00000'],
            ['Red', 'FF0000'],
            ['Orange', 'FFC000'],
            ['Yellow', 'FFFF00'],
            ['Light Green', '92D050'],
            ['Green', '00B050'],
            ['Light Blue', '00B0F0'],
            ['Blue', '0070C0'],
            ['Dark Blue', '002060'],
            ['Purple', '7030A0'],
          ] as const
        ).map(([label, value]) => ({
          label,
          action: () => void perform('table-fill', { tableFill: value }),
        })),
        null,
        { label: 'No Fill', action: () => void perform('table-fill', { tableFill: null }) },
      ],
      element,
    );
    return;
  }
  if (name === 'table') {
    if (!slideModel() || !(element instanceof HTMLElement)) return;
    openTablePicker(element, async (table) => {
      const current = model();
      if (!current || !slideModel()) return false;
      const w = Math.round(current.width * 0.8);
      const h = Math.round(Math.min(current.height * 0.8, table.rows * 365760));
      return perform('insert', {
        preset: 'table',
        table,
        changes: {
          bounds: {
            x: Math.round((current.width - w) / 2),
            y: Math.round((current.height - h) / 2),
            w,
            h,
          },
        },
      });
    });
    return;
  }
  if (name === 'file-menu') {
    menuAt([{ label: 'Page Setup…', action: () => void action('page-setup', element) }], element);
    return;
  }
  if (name === 'slide-size' || name === 'page-setup') {
    const current = model();
    if (!current) return;
    const revision = office.getState().revision;
    const save = (size: NonNullable<EditCommand['size']>, scaleContent: boolean) =>
      send(command('slide-size', { size, scaleContent }), undefined, revision);
    const setup = () =>
      openPageSetup(
        { ...current, ...(current.sizeType === undefined ? {} : { type: current.sizeType }) },
        save,
      );
    if (name === 'page-setup') {
      setup();
      return;
    }
    menuAt(
      [
        ...slideSizePresets.map(({ label, size }) => ({
          label,
          checked: current.width === size.width && current.height === size.height,
          action: () => {
            if (current.width !== size.width || current.height !== size.height)
              confirmSlideSize(size, save);
          },
        })),
        null,
        {
          label: 'Page Setup…',
          action: setup,
        },
      ],
      element,
    );
    return;
  }
  if (name === 'picture' || name === 'picture-replace') {
    menuAt(pictureSourceMenu(name === 'picture-replace'), element);
    return;
  }
  if (name === 'strike') {
    await perform('update', {
      changes: { format: { strike: !shape?.format.strike || shape.format.strike === 'noStrike' } },
    });
    return;
  }
  if (name === 'superscript' || name === 'subscript') {
    const offset = name === 'superscript' ? 0.3 : -0.25;
    await perform('update', {
      changes: { format: { baseline: shape?.format.baseline === offset ? 0 : offset } },
    });
    return;
  }
  if (name === 'bullets' || name === 'numbering') {
    const style = name === 'bullets' ? 'bullet' : 'number';
    await perform('update', { changes: { bullets: shape?.bullets === style ? 'none' : style } });
    return;
  }
  if (name === 'promote' || name === 'demote') {
    await perform('update', { changes: { indent: name === 'promote' ? -1 : 1 } });
    return;
  }
  const choices: Record<string, Array<[string, NonNullable<EditCommand['changes']>]>> = {
    'character-spacing': [
      ['Very Tight', { format: { spc: -300 } }],
      ['Tight', { format: { spc: -150 } }],
      ['Normal', { format: { spc: 0 } }],
      ['Loose', { format: { spc: 300 } }],
      ['Very Loose', { format: { spc: 600 } }],
    ],
    highlight: [
      ['Yellow', { format: { highlight: 'FFFF00' } }],
      ['Green', { format: { highlight: '00FF00' } }],
      ['Cyan', { format: { highlight: '00FFFF' } }],
      ['No Color', { format: { highlight: null } }],
    ],
    'line-spacing': [1, 1.5, 2, 2.5, 3].map((value) => [String(value), { lineSpacing: value }]),
    columns: [
      ['One Column', { columns: 1 }],
      ['Two Columns', { columns: 2 }],
      ['Three Columns', { columns: 3 }],
    ],
    'text-direction': [
      ['Horizontal', { direction: 'horz' }],
      ['Rotate all text 90°', { direction: 'vert' }],
      ['Rotate all text 270°', { direction: 'vert270' }],
      ['Stacked', { direction: 'wordArtVert' }],
    ],
    'text-anchor': [
      ['Top', { anchor: 'top' }],
      ['Middle', { anchor: 'center' }],
      ['Bottom', { anchor: 'bottom' }],
    ],
  };
  if (choices[name]) {
    menuAt(
      [
        ...choices[name].map(([label, changes]) => ({
          label,
          action: () => void perform('update', { changes }),
        })),
        ...(name === 'character-spacing'
          ? [
              {
                label: 'More Spacing...',
                action: () => void action('font-spacing-dialog', element),
              },
            ]
          : []),
        ...(name === 'line-spacing'
          ? [
              {
                label: 'Line Spacing Options...',
                action: () => void action('paragraph-dialog', element),
              },
            ]
          : []),
        ...(name === 'columns'
          ? [{ label: 'More Columns...', action: () => shapeProperties.showColumns() }]
          : []),
      ],
      element,
    );
    return;
  }
  if (name === 'change-case') {
    menuAt(
      (
        [
          ['Sentence case.', 'sentence'],
          ['lowercase', 'lower'],
          ['UPPERCASE', 'upper'],
          ['Capitalize Each Word', 'title'],
          ['tOGGLE cASE', 'toggle'],
        ] as const
      ).map(([label, mode]) => ({
        label,
        action: () =>
          void perform('update', {
            changes: {
              changeCase: mode,
            },
          }),
      })),
      element,
    );
    return;
  }
  if (name === 'undo' || name === 'redo') {
    await send(undefined, name);
    return;
  }
  if (name === 'section-menu') {
    menuAt(sectionMenu(), element);
    return;
  }
  if (name === 'slide-add-layout' || name === 'slide-layout') {
    const deck = model();
    if (!deck) return;
    const rect = element.getBoundingClientRect();
    office.menu(
      deck.layouts.map((layout) => ({
        label: layout.name || 'Untitled Layout',
        thumbnail: {
          width: deck.width,
          height: deck.height,
          boxes: layout.placeholders.flatMap((p) => (p.bounds ? [p.bounds] : [])),
        },
        ...(name === 'slide-layout' ? { checked: slideModel()?.layout === layout.key } : {}),
        action: () =>
          void perform(name === 'slide-layout' ? 'slide-layout' : 'slide-add', {
            layout: layout.key,
          }),
      })),
      rect.left,
      rect.bottom + 3,
      element,
    );
    return;
  }
  if (name.startsWith('slide-')) {
    await perform(name as EditCommand['type']);
    return;
  }
  if (name === 'bold' || name === 'italic') {
    await perform('update', { changes: { format: { [name]: !shape?.format[name] } } });
    return;
  }
  if (name === 'underline') {
    await perform('update', {
      changes: {
        format: {
          underline: !shape?.format.underline || shape.format.underline === 'none' ? 'sng' : 'none',
        },
      },
    });
    return;
  }
  if (name.startsWith('align-')) {
    await perform('update', {
      changes: {
        align: name.slice(6) as NonNullable<NonNullable<EditCommand['changes']>['align']>,
      },
    });
    return;
  }
  if (name.startsWith('font-')) {
    await perform('update', {
      changes: {
        format: {
          size: Math.min(
            400,
            Math.max(1, (shape?.format.size ?? 18) + (name === 'font-grow' ? 2 : -2)),
          ),
        },
      },
    });
    return;
  }
  if (name === 'text') {
    tool = 'text';
    selected.clear();
    render();
    focusCanvas();
    return;
  }
  if (name === 'picture-crop') {
    if (selection().length !== 1 || shape?.kind !== 'picture') return;
    cropping = cropping === shape.id ? null : shape.id;
    cropAspectRatio = undefined;
    render();
    focusCanvas();
    return;
  }
  if (name === 'picture-corrections') {
    const src = shape && elementFor(shape.id)?.querySelector('image')?.getAttribute('href');
    menuAt(
      [
        ...[-40, -20, 0, 20, 40].flatMap((contrast) =>
          [-40, -20, 0, 20, 40].map((brightness) => ({
            label: `Brightness: ${brightness}%, Contrast: ${contrast}%`,
            preview: () =>
              pictureMenuPreview({
                imageBrightness: brightness / 100,
                imageContrast: contrast / 100,
              }),
            ...(src
              ? {
                  picture: {
                    src,
                    opacity: shape?.imageOpacity ?? 1,
                    brightness: brightness / 100,
                    contrast: contrast / 100,
                  },
                }
              : {}),
            action: () => {
              void perform('update', {
                changes: { imageBrightness: brightness / 100, imageContrast: contrast / 100 },
              });
            },
          })),
        ),
        { label: 'Picture Corrections Options...', action: () => shapeProperties.showPicture() },
      ],
      element,
    );
    return;
  }
  if (name === 'picture-transparency') {
    const src = shape && elementFor(shape.id)?.querySelector('image')?.getAttribute('href');
    menuAt(
      [
        ...[0, 15, 30, 50, 65, 80, 95].map((percent) => ({
          label: `${percent}%`,
          preview: () => pictureMenuPreview({ imageOpacity: 1 - percent / 100 }),
          ...(src ? { picture: { src, opacity: 1 - percent / 100 } } : {}),
          action: () => {
            void perform('update', { changes: { imageOpacity: 1 - percent / 100 } });
          },
        })),
        { label: 'Picture Transparency Options...', action: () => shapeProperties.showPicture() },
      ],
      element,
    );
    return;
  }
  if (name === 'picture-crop-menu') {
    const applyCrop = async (changes: NonNullable<EditCommand['changes']>, ratio?: number) => {
      const ids = [...selected],
        slide = office.getState().index;
      const applied = await send({ type: 'update', slide, ids, changes });
      if (
        applied &&
        ids.length === 1 &&
        selected.size === 1 &&
        selected.has(ids[0]!) &&
        office.getState().index === slide
      ) {
        cropping = ids[0]!;
        cropAspectRatio = ratio;
        render();
        focusCanvas();
      }
    };
    const cropRatioItem = (label: string) => ({
      label,
      action: () => {
        const [width, height] = label.split(':').map(Number);
        const ratio = width! / height!;
        void applyCrop({ imageCropAspectRatio: ratio }, ratio);
      },
    });
    menuAt(
      [
        {
          label: 'Crop',
          disabled: selection().length !== 1,
          action: () => {
            cropAspectRatio = undefined;
            cropping = selection()[0]?.id ?? null;
            render();
            focusCanvas();
          },
        },
        {
          label: 'Crop to Shape',
          children: shapeCategories
            .filter((category) => category.label !== 'Action Buttons')
            .flatMap((category) => [
              { label: category.label, heading: true },
              ...category.shapes.map(({ preset, label }) => ({
                label,
                icon: getPresetShapePath(preset)!,
                action: () => {
                  void perform('update', { changes: { imageCropShape: preset } });
                },
              })),
            ]),
        },
        {
          label: 'Aspect Ratio',
          children: [
            { label: 'Square', heading: true },
            ...['1:1'].map(cropRatioItem),
            { label: 'Portrait', heading: true },
            ...['2:3', '3:4', '3:5', '4:5', '5:8', '9:16'].map(cropRatioItem),
            { label: 'Landscape', heading: true },
            ...['3:2', '4:3', '5:3', '5:4', '8:5', '16:9'].map(cropRatioItem),
          ],
        },
        {
          label: 'Fill',
          action: () => {
            void applyCrop({ imageFit: 'fill' });
          },
        },
        {
          label: 'Fit',
          action: () => {
            void applyCrop({ imageFit: 'fit' });
          },
        },
      ],
      element,
    );
    return;
  }
  if (name === 'shapes') {
    menuAt(
      [
        { label: 'Lines', heading: true },
        ...lineTools.map((line) => ({
          label: line.label,
          icon: linePath(line.preset),
          action: () => {
            tool = line.preset;
            selected.clear();
            render();
            focusCanvas();
          },
        })),
        ...shapeCategories.flatMap((category) => [
          { label: category.label, heading: true },
          ...category.shapes.map(({ preset, label }) => ({
            label,
            icon: getPresetShapePath(preset)!,
            action: () => {
              tool = preset;
              selected.clear();
              render();
              focusCanvas();
            },
          })),
        ]),
      ],
      element,
    );
    return;
  }
  if (name === 'arrange' || name === 'arrange-menu') {
    menuAt(arrangeMenu(name === 'arrange-menu'), element);
    return;
  }
  if (name === 'select-all') {
    selected.clear();
    for (const shape of slideModel()?.shapes ?? []) selected.add(shape.id);
    render();
    focusCanvas();
    return;
  }
  if (name === 'selection') {
    byId('selection-pane').hidden = !byId('selection-pane').hidden;
    if (!byId('selection-pane').hidden) byId('animation-pane').hidden = true;
    renderSelection();
    return;
  }
  if (name === 'notes') {
    byId('notes-pane').hidden = !byId('notes-pane').hidden;
    if (!byId('notes-pane').hidden) byId('speaker-notes').focus();
    return;
  }
  if (name === 'sorter') {
    setDocumentView(true);
    return;
  }
  if (name === 'copy' || name === 'cut') {
    if (clipboardPending || !selected.size) return;
    const state = office.getState();
    const slide = state.index;
    const ids = [...selected];
    const key = slideModel()!.key;
    const revision = state.revision;
    clipboardPending = true;
    updateControls();
    try {
      error('');
      const response = await fetch('/edit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ revision, snapshot: true }),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error ?? 'Could not copy objects.');
      if (typeof result.snapshot !== 'string' || result.snapshot.length > 26_000_000)
        throw new Error('The clipboard snapshot is too large.');
      const next = { slide, ids, snapshot: result.snapshot, key, cut: name === 'cut' };
      if (name === 'cut' && !(await send({ type: 'delete', slide, ids }, undefined, revision)))
        return;
      clipboard = next;
    } catch (cause) {
      error(cause instanceof Error ? cause.message : String(cause));
    } finally {
      clipboardPending = false;
      updateControls();
    }
    return;
  }
  if (name === 'paste' && clipboard && !clipboardPending) {
    await perform('paste', {
      source: clipboard.slide,
      ids: clipboard.ids,
      snapshot: clipboard.snapshot,
      pasteSlideCoordinates: true,
      pasteOffset: !clipboard.cut && slideModel()?.key === clipboard.key,
    });
  }
}
async function saveTransitionSound(value: EditorTransitionSound, key = slideModel()?.key) {
  if (busy || pendingTransitionDuration || pendingAdvanceTiming || !key) return;
  pendingSound = { key, value };
  updateControls();
  try {
    if (!(await finishText()) || slideModel()?.key !== key) return;
    await perform('transition-sound', { transitionSound: value });
  } finally {
    pendingSound = null;
    updateControls();
  }
}
byId<HTMLSelectElement>('transition-sound').onchange = async (event) => {
  const value = (event.target as HTMLSelectElement).value;
  if (value === 'other') {
    soundFileSlide = slideModel()?.key;
    const file = byId<HTMLInputElement>('transition-sound-file');
    file.value = '';
    file.click();
    updateControls();
  } else if (value === 'none' || value === 'stop') {
    await saveTransitionSound({ kind: value });
    updateControls();
  }
};
byId<HTMLInputElement>('transition-sound-loop').onchange = async (event) => {
  const value = slideModel()?.transitionSound;
  if (value?.kind === 'play')
    await saveTransitionSound({ ...value, loop: (event.target as HTMLInputElement).checked });
  updateControls();
};
byId<HTMLInputElement>('transition-sound-file').onchange = async (event) => {
  const file = (event.target as HTMLInputElement).files?.[0];
  const key = soundFileSlide;
  if (!file || loadingTransitionSound || !key) return;
  loadingTransitionSound = true;
  updateControls();
  try {
    if (file.size > 20_000_000) throw new Error('Choose a WAV sound under 20 MB.');
    const bytes = new Uint8Array(await file.arrayBuffer());
    const decoder = new TextDecoder();
    if (
      decoder.decode(bytes.slice(0, 4)) !== 'RIFF' ||
      decoder.decode(bytes.slice(8, 12)) !== 'WAVE'
    )
      throw new Error('Choose a WAV sound.');
    let binary = '';
    for (let offset = 0; offset < bytes.length; offset += 8192)
      binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
    await saveTransitionSound(
      { kind: 'play', name: file.name, base64: btoa(binary), loop: false },
      key,
    );
  } catch (cause) {
    error(cause instanceof Error ? cause.message : String(cause));
  } finally {
    loadingTransitionSound = false;
    updateControls();
  }
};
byId('animation-pane-close').onclick = () => {
  byId('animation-pane').hidden = true;
  updateControls();
  byId('animation-pane-toggle').focus();
};
async function changeAnimationStart(start: 'clickEffect' | 'withEffect' | 'afterEffect') {
  if (busy || pendingAnimationTiming || !selectedAnimation) return;
  const target = selectedAnimation;
  const animationIds = selectedAnimationIds();
  pendingAnimationTiming = true;
  try {
    if (!(await finishText()) || slideModel()?.key !== target.slide) return;
    await perform('animation-settings', { animationIds, animationSettings: { start } });
  } finally {
    pendingAnimationTiming = false;
    updateControls();
  }
}
byId<HTMLSelectElement>('animation-start').onchange = (event) => {
  void changeAnimationStart(
    (event.target as HTMLSelectElement).value as 'clickEffect' | 'withEffect' | 'afterEffect',
  );
};
for (const kind of ['duration', 'delay'] as const) {
  const input = byId<HTMLInputElement>('animation-' + kind);
  input.oninput = () => input.setCustomValidity('');
  input.onchange = async () => {
    const milliseconds = input.value.trim() ? Math.round(Number(input.value) * 1000) : NaN;
    if (!Number.isFinite(milliseconds) || milliseconds < 0 || milliseconds > 4294967295) {
      input.setCustomValidity('Enter a time in seconds.');
      input.reportValidity();
      return;
    }
    input.setCustomValidity('');
    if (busy || pendingAnimationTiming || !selectedAnimation) return;
    const target = selectedAnimation;
    const animationIds = selectedAnimationIds();
    pendingAnimationTiming = true;
    try {
      if (!(await finishText()) || slideModel()?.key !== target.slide) return;
      await perform('animation-settings', {
        animationIds,
        animationSettings:
          kind === 'duration' ? { durationMs: milliseconds } : { delayMs: milliseconds },
      });
    } finally {
      pendingAnimationTiming = false;
      updateControls();
    }
  };
}
byId<HTMLInputElement>('transition-duration').onchange = async (event) => {
  const input = event.target as HTMLInputElement;
  const milliseconds = input.value.trim() ? Math.round(Number(input.value) * 1000) : NaN;
  if (!Number.isFinite(milliseconds) || milliseconds < 0 || milliseconds > 4294967295) {
    input.setCustomValidity('Enter a duration in seconds.');
    input.reportValidity();
    return;
  }
  input.setCustomValidity('');
  if (busy || pendingTransitionDuration || pendingAdvanceTiming) return;
  const key = slideModel()?.key;
  pendingTransitionDuration = true;
  try {
    if (!(await finishText()) || slideModel()?.key !== key) return;
    await perform('transition-duration', { transitionDurationMs: milliseconds });
  } finally {
    pendingTransitionDuration = false;
    updateControls();
  }
};
byId<HTMLInputElement>('transition-duration').oninput = (event) =>
  (event.target as HTMLInputElement).setCustomValidity('');
for (const id of ['advance-click', 'advance-after', 'advance-time']) {
  byId<HTMLInputElement>(id).onchange = async () => {
    const click = byId<HTMLInputElement>('advance-click').checked;
    const after = byId<HTMLInputElement>('advance-after').checked;
    const input = byId<HTMLInputElement>('advance-time');
    const match = input.value.trim().match(/^(?:(\d+):)?(\d+(?:\.\d+)?)$/);
    const seconds = match ? Number(match[1] ?? 0) * 60 + Number(match[2]) : NaN;
    const delay = Math.round(seconds * 1000);
    if (
      after &&
      (!Number.isFinite(delay) ||
        delay < 0 ||
        delay > 4294967295 ||
        (match?.[1] !== undefined && Number(match[2]) >= 60))
    ) {
      input.setCustomValidity('Enter a time in minutes:seconds, for example 00:05.00.');
      input.reportValidity();
      return;
    }
    input.setCustomValidity('');
    if (busy || pendingTransitionDuration || pendingAdvanceTiming) {
      updateControls();
      return;
    }
    const key = slideModel()?.key;
    const advanceTiming = { advanceOnClick: click, advanceAfterMs: after ? delay : null };
    pendingAdvanceTiming = { slide: key, ...advanceTiming };
    try {
      if (!(await finishText()) || slideModel()?.key !== key) return;
      await perform('slide-advance', { advanceTiming });
    } finally {
      pendingAdvanceTiming = null;
      updateControls();
    }
  };
}
byId<HTMLInputElement>('advance-time').oninput = (event) =>
  (event.target as HTMLInputElement).setCustomValidity('');
for (const button of document.querySelectorAll<HTMLElement>('[data-edit]'))
  button.addEventListener('click', () => void action(button.dataset.edit!, button));
for (const input of document.querySelectorAll<HTMLInputElement>('[data-table-style]'))
  input.addEventListener('change', () => {
    const shape = activeShape();
    if (!shape?.table || pendingTableStyle) return;
    const key = input.dataset.tableStyle!;
    pendingTableStyle = { slide: slideModel()?.key, id: shape.id, key, value: input.checked };
    void perform('table-style', { tableStyle: { [key]: input.checked } }).finally(() => {
      pendingTableStyle = null;
      updateControls();
    });
  });
for (const axis of ['rows', 'columns'] as const) {
  const input = byId<HTMLInputElement>('table-size-' + axis);
  input.onchange = () => {
    if (!input.value || !input.checkValidity()) {
      updateControls();
      return;
    }
    void perform('table-cell-size', {
      tableAxis: axis,
      tableCellSize: Math.round(Number(input.value) * 914400),
    });
  };
}
byId<HTMLInputElement>('font-name').onchange = () =>
  void perform('update', {
    changes: { format: { font: byId<HTMLInputElement>('font-name').value } },
  });
byId<HTMLInputElement>('font-size').onchange = () =>
  void perform('update', {
    changes: { format: { size: Number(byId<HTMLInputElement>('font-size').value) } },
  });
byId<HTMLInputElement>('font-color').onchange = () =>
  void perform('update', {
    changes: { format: { color: byId<HTMLInputElement>('font-color').value } },
  });
byId<HTMLInputElement>('shape-fill').onchange = () =>
  void perform('update', { changes: { fill: byId<HTMLInputElement>('shape-fill').value } });
byId<HTMLTextAreaElement>('speaker-notes').onchange = () =>
  void send(command('notes', { text: byId<HTMLTextAreaElement>('speaker-notes').value }));
function setDocumentView(sorter: boolean) {
  document.body.classList.toggle('slide-sorter', sorter);
  for (const id of ['view-normal', 'status-normal'])
    byId(id).setAttribute('aria-pressed', String(!sorter));
  byId('view-sorter').setAttribute('aria-pressed', String(sorter));
  if (!sorter) focusCanvas();
}
for (const id of ['view-normal', 'status-normal'])
  byId(id).addEventListener('click', () => setDocumentView(false));
byId('thumbnails').addEventListener('dblclick', (event) => {
  if ((event.target as HTMLElement).closest('.section-heading')) return;
  setDocumentView(false);
  focusCanvas();
});
const collapsedSections = new Set<number>();
let selectedSection: number | null = null;
function focusSection(index: number) {
  selectedSection = index;
  renderSections();
  document.querySelector<HTMLElement>('.section-heading[data-section="' + index + '"]')?.focus();
}
document.addEventListener('focusin', (event) => {
  if ((event.target as HTMLElement).closest('.thumbnail') && selectedSection !== null) {
    selectedSection = null;
    renderSections();
  }
});
function currentSection() {
  if (selectedSection !== null && model()?.sections[selectedSection]) return selectedSection;
  return (
    model()?.sections.findIndex((section) => section.slides.includes(slideModel()?.key ?? '')) ?? -1
  );
}
function renameSection(section: number) {
  const value = model()?.sections[section];
  if (!value) return;
  const dialog = document.createElement('dialog');
  dialog.className = 'section-dialog';
  dialog.setAttribute('aria-label', 'Rename Section');
  const form = document.createElement('form');
  const heading = document.createElement('h2');
  heading.textContent = 'Rename Section';
  const label = document.createElement('label');
  label.textContent = 'Section name:';
  const input = document.createElement('input');
  input.value = value.name;
  input.required = true;
  label.append(input);
  const buttons = document.createElement('div');
  const cancel = document.createElement('button');
  cancel.type = 'button';
  cancel.textContent = 'Cancel';
  cancel.onclick = () => dialog.close();
  const submit = document.createElement('button');
  submit.type = 'submit';
  submit.textContent = 'Rename';
  buttons.append(cancel, submit);
  form.append(heading, label, buttons);
  dialog.append(form);
  form.onsubmit = (event) => {
    event.preventDefault();
    void perform('section-rename', { section, text: input.value }).then((ok) => {
      if (ok) dialog.close();
    });
  };
  dialog.onclose = () => dialog.remove();
  document.body.append(dialog);
  dialog.showModal();
  input.select();
}
async function addSection() {
  selectedSection = null;
  if (await perform('section-add', { text: 'Untitled Section' })) renameSection(currentSection());
}
async function moveSection(section: number, to: number) {
  const key = slideModel()?.key;
  const previous = new Set(collapsedSections);
  if (!(await perform('section-move', { section, to }))) return;
  collapsedSections.clear();
  const order = (model()?.sections ?? []).map((_, i) => i);
  order.splice(to, 0, order.splice(section, 1)[0]!);
  order.forEach((oldIndex, newIndex) => {
    if (previous.has(oldIndex)) collapsedSections.add(newIndex);
  });
  const position = model()?.slides.findIndex((slide) => slide.key === key) ?? -1;
  if (position >= 0) office.selectSlide(position, false);
  focusSection(to);
}
function sectionMenu(section = currentSection()): MenuItem[] {
  return [
    { label: 'Add Section', action: () => void addSection() },
    { label: 'Rename Section', disabled: section < 0, action: () => renameSection(section) },
    {
      label: 'Remove Section',
      disabled: section < 0,
      action: () => void perform('section-remove', { section }),
    },
    {
      label: 'Remove Section & Slides',
      disabled: section < 0,
      action: () => void perform('section-delete', { section }),
    },
    {
      label: 'Remove All Sections',
      disabled: !model()?.sections.length,
      action: () => void perform('section-remove-all'),
    },
    null,
    {
      label: 'Move Section Up',
      disabled: section <= 0,
      action: () => void moveSection(section, section - 1),
    },
    {
      label: 'Move Section Down',
      disabled: section < 0 || section >= (model()?.sections.length ?? 0) - 1,
      action: () => void moveSection(section, section + 1),
    },
    null,
    {
      label: 'Collapse All',
      disabled: !model()?.sections.length,
      action: () => {
        model()?.sections.forEach((_, i) => collapsedSections.add(i));
        renderSections();
      },
    },
    {
      label: 'Expand All',
      disabled: !model()?.sections.length,
      action: () => {
        collapsedSections.clear();
        renderSections();
      },
    },
  ];
}
function renderSections() {
  const focusedSection = (document.activeElement as HTMLElement | null)?.dataset.section;
  const deck = model();
  const items = [...byId('thumbnails').children] as HTMLElement[];
  for (const item of items) {
    item.querySelectorAll('.section-heading').forEach((header) => header.remove());
    item.hidden = false;
    item.classList.remove('has-section');
    const thumbnail = item.querySelector<HTMLElement>('.thumbnail');
    if (thumbnail) thumbnail.hidden = false;
  }
  byId('thumbnails').parentElement?.querySelector('.empty-section-list')?.remove();
  if (!deck) return;
  const emptySections = document.createElement('div');
  emptySections.className = 'empty-section-list';
  if (!items.length && deck.sections.length) byId('thumbnails').after(emptySections);
  for (const [index, section] of deck.sections.entries()) {
    const start = deck.slides.findIndex((slide) => section.slides.includes(slide.key));
    const next = deck.sections.slice(index + 1).flatMap((s) => s.slides)[0];
    const trailing = start < 0 && next === undefined;
    const anchor = trailing
      ? items.length - 1
      : start >= 0
        ? start
        : Math.max(
            0,
            deck.slides.findIndex((slide) => slide.key === next),
          );
    const item = items[anchor] ?? emptySections;
    const header = document.createElement('button');
    header.className = 'section-heading';
    header.dataset.section = String(index);
    if (trailing) header.style.order = '1';
    header.draggable = true;
    header.ondragstart = (event) => {
      if (busy) {
        event.preventDefault();
        return;
      }
      draggedSection = index;
      event.dataTransfer!.setData('text/plain', section.name);
      event.dataTransfer!.effectAllowed = 'move';
    };
    header.ondragover = (event) => {
      if (draggedSection === null) return;
      event.preventDefault();
      header.classList.add('section-drop-target');
      header.classList.toggle('section-drop-after', draggedSection < index);
    };
    header.ondragleave = () => header.classList.remove('section-drop-target');
    header.ondragend = () => {
      draggedSection = null;
      document
        .querySelectorAll('.section-drop-target')
        .forEach((node) => node.classList.remove('section-drop-target'));
    };
    header.ondrop = (event) => {
      if (draggedSection === null) return;
      event.preventDefault();
      event.stopPropagation();
      header.classList.remove('section-drop-target');
      if (draggedSection !== index) void moveSection(draggedSection, index);
      draggedSection = null;
    };
    const disclosure = document.createElement('span');
    disclosure.className = 'section-disclosure';
    disclosure.textContent = collapsedSections.has(index) ? '▸' : '▾';
    disclosure.setAttribute('aria-hidden', 'true');
    header.append(
      disclosure,
      document.createTextNode(section.name + ' (' + section.slides.length + ')'),
    );
    header.classList.toggle('section-selected', selectedSection === index);
    header.setAttribute('aria-expanded', String(!collapsedSections.has(index)));
    header.onclick = (event) => {
      if ((event.target as HTMLElement).closest('.section-disclosure')) {
        if (collapsedSections.has(index)) collapsedSections.delete(index);
        else collapsedSections.add(index);
      }
      focusSection(index);
    };
    header.oncontextmenu = (event) => {
      event.preventDefault();
      event.stopPropagation();
      focusSection(index);
      const anchor = document.querySelector<HTMLElement>(
        '.section-heading[data-section="' + index + '"]',
      )!;
      office.menu(sectionMenu(index), event.clientX, event.clientY, anchor);
    };
    item.append(header);
    if (focusedSection === String(index)) header.focus({ preventScroll: true });
    item.classList.add('has-section');
    if (collapsedSections.has(index))
      for (const [i, slide] of deck.slides.entries()) {
        if (section.slides.includes(slide.key)) {
          const thumbnail = items[i]?.querySelector<HTMLElement>('.thumbnail');
          if (thumbnail) thumbnail.hidden = true;
        }
      }
  }
  for (const item of items)
    item.hidden =
      !!item.querySelector<HTMLElement>('.thumbnail')?.hidden &&
      !item.classList.contains('has-section');
}
function slideMenu(): MenuItem[] {
  return [
    { label: 'New Slide', action: () => void perform('slide-add') },
    { label: 'Add Section', action: () => void addSection() },
    { label: 'Duplicate Slide', action: () => void perform('slide-duplicate') },
    { label: 'Delete Slide', action: () => void perform('slide-delete') },
    null,
    {
      label: slideModel()?.hidden ? 'Unhide Slide' : 'Hide Slide',
      action: () => void perform('slide-hidden', { hidden: !slideModel()?.hidden }),
    },
    null,
    { label: 'Zoom…', action: () => byId('zoom-level').click() },
    { label: 'Slide Show', action: () => byId('present').click() },
  ];
}
byId('thumbnails').addEventListener(
  'contextmenu',
  (event) => {
    if (office.getState().presenting) return;
    const thumbnail = (event.target as HTMLElement).closest('.thumbnail');
    if (!thumbnail) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    office.selectSlide([...byId('thumbnails').children].indexOf(thumbnail.parentElement!), true);
    office.menu(slideMenu(), event.clientX, event.clientY, thumbnail);
  },
  true,
);
async function selectTablePart(part: 'table' | 'column' | 'row') {
  if (busy) return;
  const shape = selection()[0];
  if (!shape?.table) return;
  const target = textEdit?.cell ?? cellSelection?.range;
  if (!target && part !== 'table') return;
  const slide = slideModel()?.key;
  const span = target && shape.table.cells[target.row]?.[target.column]?.span;
  const range = expandTableSelection(shape.table, {
    row: part === 'row' ? target!.row : 0,
    column: part === 'column' ? target!.column : 0,
    rows:
      part === 'row'
        ? (cellSelection?.range.rows ?? span?.rowSpan ?? 1)
        : shape.table.rowHeights.length,
    columns:
      part === 'column'
        ? (cellSelection?.range.columns ?? span?.gridSpan ?? 1)
        : shape.table.columnWidths.length,
  });
  if (!(await finishText()) || slideModel()?.key !== slide) return;
  selected.clear();
  selected.add(shape.id);
  cellSelection = {
    id: shape.id,
    slide: office.getState().index,
    revision: office.getState().revision,
    range,
  };
  render();
  focusCanvas();
}

function tableSelectMenu(): MenuItem[] {
  const table = selection()[0]?.table;
  const range = cellSelection?.range;
  const allRows = !!range && range.row === 0 && range.rows === table?.rowHeights.length;
  const allColumns = !!range && range.column === 0 && range.columns === table?.columnWidths.length;
  return [
    {
      label: 'Select Table',
      disabled: allRows && allColumns,
      action: () => void selectTablePart('table'),
    },
    { label: 'Select Column', disabled: allRows, action: () => void selectTablePart('column') },
    { label: 'Select Row', disabled: allColumns, action: () => void selectTablePart('row') },
  ];
}

function textClipboardMenu(): MenuItem[] {
  const hasRange = !!textEdit && textEdit.field.selectionStart !== textEdit.field.selectionEnd;
  return (['cut', 'copy', 'paste'] as const).map((name) => ({
    label: { cut: 'Cut', copy: 'Copy', paste: 'Paste' }[name],
    disabled: clipboardPending || (name === 'paste' ? !navigator.clipboard?.readText : !hasRange),
    action: () => void action(name, host),
  }));
}

function tableContextMenu(): MenuItem[] {
  const command = (label: string, name: string): NonNullable<MenuItem> => ({
    label,
    disabled:
      busy || !!document.querySelector<HTMLButtonElement>(`[data-edit="${name}"]`)?.disabled,
    action: () => void action(name, host),
  });
  return [
    {
      label: 'Insert',
      children: [
        command('Insert Columns to the Left', 'table-insert-left'),
        command('Insert Columns to the Right', 'table-insert-right'),
        command('Insert Rows Above', 'table-insert-above'),
        command('Insert Rows Below', 'table-insert-below'),
      ],
    },
    { label: 'Delete', disabled: busy, children: tableDeleteMenu() },
    { label: 'Select', disabled: busy, children: tableSelectMenu() },
    null,
    command('Merge Cells', 'table-merge'),
    command('Split Cells...', 'table-split'),
    null,
    command('Font...', 'font-dialog'),
    command('Paragraph...', 'paragraph-dialog'),
    ...(textEdit?.cell ? [command('Hyperlink...', 'link'), ...openHyperlinkMenu()] : []),
  ];
}

host.addEventListener('contextmenu', async (event) => {
  if (office.getState().presenting) return;
  event.preventDefault();
  event.stopPropagation();
  if (textEdit && event.composedPath().includes(textEdit.field)) {
    // Commit pending typing before resolving hyperlink offsets from the model.
    if (textEdit.edits.length) {
      const edit = textEdit;
      const range = { start: edit.field.selectionStart, end: edit.field.selectionEnd };
      const key = slideModel()?.key;
      if (!(await finishText()) || slideModel()?.key !== key) return;
      const shape = allShapes().find((shape) => shape.id === edit.shape.id);
      if (!shape) return;
      if (edit.cell) startCell(shape, edit.cell.row, edit.cell.column, range);
      else startText(shape, range);
      if (!textEdit) return;
    }
    const hasRange = textEdit.field.selectionStart !== textEdit.field.selectionEnd;
    office.menu(
      [
        ...textClipboardMenu(),
        null,
        ...(textEdit.cell
          ? tableContextMenu()
          : [
              {
                label: 'Delete',
                disabled: !hasRange,
                action: () => textEdit?.field.deleteSelection(),
              },
              { label: 'Select All', action: () => void action('select-all', host) },
              null,
              { label: 'Font...', action: () => void action('font-dialog', host) },
              { label: 'Paragraph...', action: () => void action('paragraph-dialog', host) },
              {
                label: 'Hyperlink...',
                disabled: busy,
                action: () => void action('link', host),
              },
              ...openHyperlinkMenu(),
            ]),
      ],
      event.clientX,
      event.clientY,
      textEdit.field,
    );
    return;
  }
  const hit = event
    .composedPath()
    .find((node) => node instanceof HTMLElement && node.dataset.shapeId) as HTMLElement | undefined;
  if (hit && !selected.has(Number(hit.dataset.shapeId))) {
    cellSelection = null;
    selected.clear();
    selected.add(Number(hit.dataset.shapeId));
    render();
  }
  const tableHit =
    hit && selection().find((shape) => shape.id === Number(hit.dataset.shapeId) && shape.table);
  if (tableHit && !cellSelection && !textEdit && !busy) {
    const currentHit = shadow.querySelector<HTMLElement>(`[data-shape-id="${tableHit.id}"]`);
    if (currentHit) startCellAt(tableHit, event, currentHit);
  }
  if (cellSelection || (tableHit && textEdit?.cell)) {
    office.menu(
      [...(textEdit?.cell ? [...textClipboardMenu(), null] : []), ...tableContextMenu()],
      event.clientX,
      event.clientY,
      textEdit?.field ?? host,
    );
    return;
  }
  office.menu(
    selected.size
      ? [
          { label: 'Cut', action: () => void action('cut', host) },
          { label: 'Copy', action: () => void action('copy', host) },
          { label: 'Paste', disabled: !clipboard, action: () => void action('paste', host) },
          null,
          {
            label: 'Edit Text',
            disabled: selected.size !== 1 || !selection()[0]?.textable,
            action: () => startText(selection()[0]!),
          },
          ...(selected.size === 1
            ? [
                {
                  label: activeShape()?.link ? 'Edit Hyperlink...' : 'Hyperlink...',
                  action: () => void action('link', host),
                },
                {
                  label: 'Action Settings...',
                  disabled: busy,
                  action: () => void action('action-settings', host),
                },
                ...openHyperlinkMenu(),
                ...(activeShape()?.link
                  ? [
                      {
                        label: 'Remove Hyperlink',
                        disabled: busy,
                        action: () => void perform('object-link', { link: null }),
                      },
                    ]
                  : []),
              ]
            : []),
          { label: 'Duplicate', action: () => void perform('duplicate') },
          { label: 'Delete', action: () => void clearSelectedContent() },
          null,
          ...arrangeMenu(),
          null,
          ...(selection().every((shape) => shape.kind === 'picture')
            ? [
                { label: 'Change Picture', children: pictureSourceMenu(true) },
                { label: 'Format Picture...', action: () => shapeProperties.showPicture() },
              ]
            : []),
          { label: 'Size and Position...', action: () => shapeProperties.show() },
        ]
      : [...slideMenu(), null, { label: 'Grid and Guides', children: guideMenu() }],
    event.clientX,
    event.clientY,
    host,
  );
});
let draggedSection: number | null = null;
let draggedSlide: number | null = null;
byId('thumbnails').addEventListener('pointerdown', (event) => {
  const thumbnail = (event.target as HTMLElement).closest<HTMLElement>('.thumbnail');
  if (thumbnail) thumbnail.draggable = true;
});
byId('thumbnails').addEventListener('dragstart', (event) => {
  const thumbnail = (event.target as HTMLElement).closest('.thumbnail');
  if (!thumbnail) return;
  draggedSlide = [...byId('thumbnails').children].indexOf(thumbnail.parentElement!);
  event.dataTransfer!.setData('text/plain', String(draggedSlide));
  event.dataTransfer!.effectAllowed = 'move';
});
byId('thumbnails').addEventListener('dragover', (event) => {
  if (draggedSlide === null) return;
  event.preventDefault();
});
byId('thumbnails').addEventListener('drop', (event) => {
  event.preventDefault();
  const thumbnail = (event.target as HTMLElement).closest('.thumbnail');
  if (thumbnail && draggedSlide !== null) {
    const to = [...byId('thumbnails').children].indexOf(thumbnail.parentElement!);
    void send({ type: 'slide-move', slide: draggedSlide, to }).then((ok) => {
      if (ok) office.selectSlide(to, true);
    });
  }
  draggedSlide = null;
});
byId('thumbnails').addEventListener('dragend', () => {
  draggedSection = null;
  document
    .querySelectorAll('.section-drop-target')
    .forEach((header) => header.classList.remove('section-drop-target'));
  draggedSlide = null;
});
document.addEventListener(
  'keydown',
  (event) => {
    if (office.getState().presenting || event.defaultPrevented || !byId('office-menu').hidden)
      return;
    if (office.getState().animationPreview && event.key === 'Escape') {
      event.preventDefault();
      office.stopAnimationPreview();
      return;
    }
    const target = event.composedPath()[0] as HTMLElement;
    if (
      target.closest(
        'input,textarea,select,[contenteditable],dialog,.table-picker,.arrow-gallery,.arrow-gallery-trigger,[role=separator],#animation-pane,.animation-marker',
      )
    )
      return;
    const modifier = event.metaKey || event.ctrlKey,
      key = event.key.toLowerCase();
    if (event.metaKey && !event.ctrlKey && !event.shiftKey) {
      if (!event.altKey && ['Digit1', 'Digit2'].includes(event.code)) {
        event.preventDefault();
        event.stopImmediatePropagation();
        setDocumentView(event.code === 'Digit2');
        return;
      }
      if (event.altKey && event.code === 'KeyR') {
        event.preventDefault();
        event.stopImmediatePropagation();
        byId('collapse-ribbon').click();
        return;
      }
    }
    if (drag) {
      event.preventDefault();
      event.stopImmediatePropagation();
      if (key === 'escape') void endDrag({ pointerId: drag.pointer }, true);
      return;
    }
    if (!target.closest('.thumbnail,.section-heading') && textFormatShortcut(event, target)) return;
    const thumbnailFocus = !!target.closest('.thumbnail');
    const sectionHeader = target.closest<HTMLElement>('.section-heading');
    if (sectionHeader && event.shiftKey && key === 'f10') {
      event.preventDefault();
      event.stopImmediatePropagation();
      const rect = sectionHeader.getBoundingClientRect();
      office.menu(
        sectionMenu(Number(sectionHeader.dataset.section)),
        rect.left,
        rect.bottom,
        sectionHeader,
      );
      return;
    }
    if (
      (sectionHeader || thumbnailFocus) &&
      modifier &&
      !event.altKey &&
      ['arrowup', 'arrowdown'].includes(key)
    ) {
      event.preventDefault();
      event.stopImmediatePropagation();
      const section = sectionHeader ? Number(sectionHeader.dataset.section) : null;
      const from = section ?? office.getState().index;
      const count =
        section === null ? (model()?.slides.length ?? 0) : (model()?.sections.length ?? 0);
      const to = event.shiftKey
        ? key === 'arrowup'
          ? 0
          : count - 1
        : from + (key === 'arrowup' ? -1 : 1);
      if (busy || to < 0 || to >= count || to === from) return;
      if (section !== null) void moveSection(section, to);
      else
        void send({ type: 'slide-move', slide: from, to }).then((ok) => {
          if (ok) office.selectSlide(to, true);
        });
      return;
    }
    if (sectionHeader && !modifier && ['arrowleft', 'arrowright'].includes(key)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      const index = Number(sectionHeader.dataset.section);
      if (key === 'arrowleft') collapsedSections.add(index);
      else collapsedSections.delete(index);
      focusSection(index);
      return;
    }
    if (
      (sectionHeader || thumbnailFocus) &&
      !modifier &&
      !event.altKey &&
      ['arrowup', 'arrowdown', 'home', 'end'].includes(key)
    ) {
      event.preventDefault();
      event.stopImmediatePropagation();
      const entries = [...document.querySelectorAll<HTMLElement>('.section-heading,.thumbnail')]
        .filter((element) => element.getClientRects().length)
        .sort(
          (a, b) =>
            a.getBoundingClientRect().top - b.getBoundingClientRect().top ||
            a.getBoundingClientRect().left - b.getBoundingClientRect().left,
        );
      const current = entries.indexOf(sectionHeader ?? target.closest<HTMLElement>('.thumbnail')!);
      const index =
        key === 'home'
          ? 0
          : key === 'end'
            ? entries.length - 1
            : Math.max(0, Math.min(entries.length - 1, current + (key === 'arrowup' ? -1 : 1)));
      const next = entries[index];
      if (next?.dataset.section !== undefined) focusSection(Number(next.dataset.section));
      else if (next)
        office.selectSlide([...byId('thumbnails').children].indexOf(next.parentElement!), true);
      return;
    }
    if (modifier && key === 'z') {
      event.preventDefault();
      event.stopImmediatePropagation();
      void send(undefined, event.shiftKey ? 'redo' : 'undo');
      return;
    }
    if (modifier && key === 'y') {
      event.preventDefault();
      event.stopImmediatePropagation();
      void send(undefined, 'redo');
      return;
    }
    if (modifier && (key === 'm' || (event.shiftKey && key === 'n'))) {
      event.preventDefault();
      event.stopImmediatePropagation();
      void perform('slide-add');
      return;
    }
    if (modifier && key === 'd') {
      event.preventDefault();
      event.stopImmediatePropagation();
      void perform(thumbnailFocus || event.shiftKey ? 'slide-duplicate' : 'duplicate');
      return;
    }
    if (key === 'escape') {
      event.preventDefault();
      event.stopImmediatePropagation();
      cancelDrawing();
      selected.clear();
      tool = null;
      render();
      focusCanvas();
      return;
    }
    if (thumbnailFocus && ['delete', 'backspace'].includes(key)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      void perform('slide-delete');
      return;
    }
    const canvasFocus = target === host || shadow.contains(target);
    if (!canvasFocus && !target.closest('.application-menubar,.ribbon,#selection-pane')) return;
    // Object commands also apply after dismissing an application/ribbon menu.
    // Handle them before text formatting (Shift+Command+B is stacking).
    const objectKey = /^Key[GBFJ]$/.test(event.code) ? event.code.slice(3).toLowerCase() : key;
    if (event.metaKey && !event.ctrlKey && event.altKey && !event.shiftKey && objectKey === 'j') {
      event.preventDefault();
      event.stopImmediatePropagation();
      if (!busy && selection().some((shape) => shape.regroupIds?.length)) void perform('regroup');
      return;
    }
    if (event.metaKey && !event.ctrlKey && event.altKey && objectKey === 'g') {
      event.preventDefault();
      event.stopImmediatePropagation();
      if (event.shiftKey) {
        if (selection().length && selection().every((shape) => shape.kind === 'group'))
          void perform('ungroup');
      } else if (selected.size > 1) void perform('group');
      return;
    }
    if (event.metaKey && !event.ctrlKey && event.shiftKey && ['b', 'f'].includes(objectKey)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      if (selected.size)
        void perform('order', {
          order:
            objectKey === 'b'
              ? event.altKey
                ? 'backward'
                : 'back'
              : event.altKey
                ? 'forward'
                : 'front',
        });
      return;
    }
    if (!canvasFocus) return;
    if (modifier && ['a', 'c', 'x', 'v', 'b', 'i', 'u'].includes(key)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      const names: Record<string, string> = {
        a: 'select-all',
        c: 'copy',
        x: 'cut',
        v: 'paste',
        b: 'bold',
        i: 'italic',
        u: 'underline',
      };
      void action(names[key]!, host);
      return;
    }
    if (key === 'tab') {
      const last = [...selected].at(-1);
      const parent = allShapes().find((shape) =>
        shape.children?.some((child) => child.id === last),
      );
      const shapes = (parent?.children ?? slideModel()?.shapes ?? []).filter(
        (shape) =>
          shape.parentTransform !== null && visibleShapes().some((item) => item.id === shape.id),
      );
      if (!shapes.length) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      const index = shapes.findIndex((s) => s.id === last);
      const next =
        index < 0
          ? event.shiftKey
            ? shapes.length - 1
            : 0
          : (index + (event.shiftKey ? -1 : 1) + shapes.length) % shapes.length;
      selected.clear();
      selected.add(shapes[next]!.id);
      render();
      return;
    }
    if (!selected.size) return;
    if (['delete', 'backspace'].includes(key)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      void clearSelectedContent();
      return;
    }
    if (event.altKey && !modifier && !event.shiftKey && ['arrowleft', 'arrowright'].includes(key)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      void perform('update', { changes: { rotationDelta: key === 'arrowright' ? 15 : -15 } });
      return;
    }
    if (event.shiftKey && !modifier && !event.altKey && key.startsWith('arrow')) {
      event.preventDefault();
      event.stopImmediatePropagation();
      const horizontal = key === 'arrowleft' || key === 'arrowright';
      const factor = key === 'arrowright' || key === 'arrowup' ? 1.1 : 1 / 1.1;
      void perform('update', {
        positions: selection()
          .filter((shape) => shape.bounds)
          .map((shape) => {
            const box = shape.bounds!;
            const minimum = shape.kind === 'connector' ? 0 : 1;
            const w =
              horizontal || shape.aspectRatioLocked
                ? Math.max(minimum, Math.round(box.w * factor))
                : box.w;
            const h =
              !horizontal || shape.aspectRatioLocked
                ? Math.max(minimum, Math.round(box.h * factor))
                : box.h;
            return {
              id: shape.id,
              bounds: {
                x: Math.round(box.x + (box.w - w) / 2),
                y: Math.round(box.y + (box.h - h) / 2),
                w,
                h,
              },
            };
          }),
      });
      return;
    }
    if (key.startsWith('arrow') && !event.shiftKey && !event.altKey && !event.ctrlKey) {
      event.preventDefault();
      event.stopImmediatePropagation();
      const step = 12700;
      void perform('update', {
        positions: selection()
          .filter((s) => s.bounds)
          .map((s) => ({
            id: s.id,
            bounds: {
              ...s.bounds!,
              x:
                s.bounds!.x +
                localVector(
                  s,
                  key === 'arrowright' ? step : key === 'arrowleft' ? -step : 0,
                  key === 'arrowdown' ? step : key === 'arrowup' ? -step : 0,
                ).x,
              y:
                s.bounds!.y +
                localVector(
                  s,
                  key === 'arrowright' ? step : key === 'arrowleft' ? -step : 0,
                  key === 'arrowdown' ? step : key === 'arrowup' ? -step : 0,
                ).y,
            },
          })),
      });
      return;
    }
    if (key === 'enter' || key === 'f2') {
      const shape = selection()[0];
      if (shape?.textable) {
        event.preventDefault();
        event.stopImmediatePropagation();
        startText(shape);
      }
    }
  },
  true,
);
window.addEventListener('office-update', render);
window.addEventListener('office-ribbon-change', renderAnimationMarkers);
new ResizeObserver(() => {
  if (!drag && !textEdit) render();
}).observe(host);
render();
