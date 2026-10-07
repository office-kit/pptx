import type { EditorDocument } from './document.svelte.ts';

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

/** Saves the deck as a `.pptx` download; returns the version written. */
export async function downloadPptx(doc: EditorDocument): Promise<number> {
  const version = doc.version;
  const bytes = await doc.toBytes();
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: PPTX_MIME }));
  const a = document.createElement('a');
  a.href = url;
  a.download = doc.fileName.endsWith('.pptx') ? doc.fileName : `${doc.fileName}.pptx`;
  a.click();
  URL.revokeObjectURL(url);
  return version;
}
