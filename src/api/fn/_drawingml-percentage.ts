// Transitional DrawingML ST_Percentage accepts integer thousandths of a
// percent as well as the percent-suffixed shared ST_Percentage form.
export function readDrawingmlPercentage(raw: string | null, fallback: number): number {
  const value = raw?.trim();
  if (!value) return fallback;
  const percent = value.endsWith('%');
  const number = Number(percent ? value.slice(0, -1) : value);
  return Number.isFinite(number) ? number / (percent ? 100 : 100000) : fallback;
}
