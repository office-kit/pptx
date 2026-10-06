// Insert-tab commands that need more than one library call: pick a file or a
// screen, size the object the way PowerPoint does, select it.

import {
  addSlideImage,
  addSlideMedia,
  addSlideTextBox,
  emu,
  getShapeId,
  getSlideSize,
  setParagraphAlignment,
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

/** Insert ▸ WordArt: PowerPoint's "Your text here" box in the style chosen from the gallery. */
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
