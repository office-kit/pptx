// Insert-tab commands that need more than one library call: pick a file or a
// screen, size the object the way the reference desktop app does, select it.

import {
  addSlideImage,
  addSlideMedia,
  addSlideTable,
  addSlideTextBox,
  emu,
  getShapeId,
  getSlideSize,
  inches,
  setParagraphAlignment,
  setShapeTextDirection,
  setShapeTextFormat,
  type SlideShapeData,
} from '@office-kit/pptx';
import type { EditorController } from './controller.svelte.ts';
import { applyWordArtPreset, type WordArtPreset } from './wordart-presets.ts';

const EMU_PER_INCH = 914400;
const AUDIO_ICON_EMU = EMU_PER_INCH / 2;
const WIDE = 16 / 9;

export function pickFile(accept: string): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = accept;
    input.onchange = () => resolve(input.files?.[0] ?? null);
    input.oncancel = () => resolve(null);
    input.click();
  });
}

/** A box of the given aspect, half the slide wide (or `size` square), centered. */
function centered(editor: EditorController, aspect: number, size?: number) {
  const slide = getSlideSize(editor.doc.pres);
  const width = slide?.width ?? 12192000;
  const height = slide?.height ?? 6858000;
  const w = size ?? Math.min(width / 2, (height / 2) * aspect);
  const h = size ?? w / aspect;
  return { x: emu((width - w) / 2), y: emu((height - h) / 2), w: emu(w), h: emu(h) };
}

function insert(editor: EditorController, label: string, add: () => SlideShapeData) {
  const doc = editor.doc;
  if (!doc.currentSlide) return;
  try {
    const shape = doc.transact(label, add);
    doc.selectShape(doc.selection.slideIndex, getShapeId(shape));
  } catch (error) {
    editor.toast('error', error instanceof Error ? error.message : String(error));
  }
}

/** Insert ▸ Video / Audio ▸ from a file on this computer. */
export async function insertMedia(
  editor: EditorController,
  kind: 'video' | 'audio',
  label: string,
) {
  const file = await pickFile(
    kind === 'video' ? 'video/*,.mp4,.mov,.m4v,.wmv,.avi' : 'audio/*,.mp3,.m4a,.wav,.wma',
  );
  const slide = editor.doc.currentSlide;
  if (!file || !slide) return;
  const data = new Uint8Array(await file.arrayBuffer());
  const box = kind === 'video' ? centered(editor, WIDE) : centered(editor, 1, AUDIO_ICON_EMU);
  insert(editor, label, () => addSlideMedia(slide, { kind, data, ...box, name: file.name }));
}

/** Insert ▸ Screenshot: a still of a screen or window the user shares. */
export async function insertScreenshot(editor: EditorController, label: string) {
  const slide = editor.doc.currentSlide;
  if (!slide || !navigator.mediaDevices?.getDisplayMedia) return;
  let stream: MediaStream;
  try {
    stream = await navigator.mediaDevices.getDisplayMedia({ video: true, audio: false });
  } catch {
    // The user dismissed the picker.
    return;
  }
  try {
    const video = document.createElement('video');
    video.muted = true;
    video.srcObject = stream;
    await video.play();
    const canvas = document.createElement('canvas');
    canvas.width = video.videoWidth;
    canvas.height = video.videoHeight;
    canvas.getContext('2d')!.drawImage(video, 0, 0);
    const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, 'image/png'));
    if (!blob) return;
    const data = new Uint8Array(await blob.arrayBuffer());
    const box = centered(editor, canvas.width / canvas.height);
    insert(editor, label, () => addSlideImage(slide, data, { ...box, name: 'Screenshot' }));
  } finally {
    for (const track of stream.getTracks()) track.stop();
  }
}

const WORDART_PT = 54;

/** Insert ▸ WordArt: the reference desktop app's "Your text here" box in the style chosen from the gallery. */
export function insertWordArt(
  editor: EditorController,
  label: string,
  text: string,
  preset: WordArtPreset,
) {
  const slide = editor.doc.currentSlide;
  if (!slide) return;
  const box = centered(editor, 4);
  insert(editor, label, () => {
    const shape = addSlideTextBox(slide, { ...box, text });
    setShapeTextFormat(shape, { size: WORDART_PT });
    applyWordArtPreset(shape, preset);
    setParagraphAlignment(shape, 0, 'ctr');
    return shape;
  });
}

const TABLE_ROW_EMU = inches(0.4);

/**
 * Insert ▸ Table: an empty table 80% of the slide wide, centered, with the
 * first cell selected for typing.
 */
export function insertTable(
  editor: EditorController,
  label: string,
  rows: number,
  columns: number,
  options: { readonly header: boolean; readonly banded: boolean },
) {
  const doc = editor.doc;
  const slide = doc.currentSlide;
  if (!slide) return;
  const size = getSlideSize(doc.pres);
  const width = size?.width ?? 12192000;
  const height = size?.height ?? 6858000;
  const w = Math.round(width * 0.8);
  const h = Math.round(Math.min(height * 0.7, rows * TABLE_ROW_EMU));
  const cells = Array.from({ length: rows }, () => Array.from({ length: columns }, () => ''));
  doc.transact(label, () => {
    const table = addSlideTable(slide, {
      x: emu(Math.round((width - w) / 2)),
      y: emu(Math.round((height - h) / 2)),
      w: emu(w),
      h: emu(h),
      rows: cells,
      firstRow: options.header,
      bandRow: options.banded,
    });
    doc.selectCell(doc.selection.slideIndex, getShapeId(table), 0, 0);
  });
}

const TEXT_BOX_EMU = { w: inches(4), h: inches(1) };

/**
 * Insert ▸ Text Box ▸ Draw Horizontal / Vertical Text Box. The reference desktop app draws
 * the box with the mouse; the editor drops a 4 × 1 in box (1 × 4 in when
 * vertical) in the middle of the slide. A vertical box rotates its text 90°
 * (`vert="vert"`), as the reference desktop app's Mac command writes.
 */
export function insertTextBox(
  editor: EditorController,
  label: string,
  text: string,
  vertical: boolean,
) {
  const slide = editor.doc.currentSlide;
  if (!slide) return;
  const size = getSlideSize(editor.doc.pres);
  const width = size?.width ?? 12192000;
  const height = size?.height ?? 6858000;
  const w = vertical ? TEXT_BOX_EMU.h : TEXT_BOX_EMU.w;
  const h = vertical ? TEXT_BOX_EMU.w : TEXT_BOX_EMU.h;
  insert(editor, label, () => {
    const shape = addSlideTextBox(slide, {
      x: emu(Math.round((width - w) / 2)),
      y: emu(Math.round((height - h) / 2)),
      w: emu(w),
      h: emu(h),
      text,
    });
    if (vertical) setShapeTextDirection(shape, 'vert');
    return shape;
  });
}
