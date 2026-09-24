// Keep editing overlays consistent with the slide renderer when a document's
// font is not installed on the computer displaying the preview.
export function textFontFamily(font?: string | null): string {
  const fallback = "Calibri, 'Helvetica Neue', Arial, sans-serif";
  return font ? `${font}, ${fallback}` : fallback;
}
