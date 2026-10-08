import type { EditorDocument } from './document.svelte.ts';

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

/** Hands `.pptx` bytes to the browser as a download. */
export function downloadPptxBytes(bytes: Uint8Array, fileName: string): void {
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: PPTX_MIME }));
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName.endsWith('.pptx') ? fileName : `${fileName}.pptx`;
  a.click();
  URL.revokeObjectURL(url);
}

/** Saves the deck as a `.pptx` download; returns the version written. */
export async function downloadPptx(doc: EditorDocument): Promise<number> {
  const version = doc.version;
  downloadPptxBytes(await doc.toBytes(), doc.fileName);
  return version;
}
