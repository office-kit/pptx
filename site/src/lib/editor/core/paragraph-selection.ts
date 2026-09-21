/** Paragraphs touched by a UTF-16 selection; the end offset is exclusive. */
export function paragraphsInTextRange(
  lengths: readonly number[],
  range: { start: number; end: number },
): number[] {
  const selected: number[] = [];
  let offset = 0;
  for (let index = 0; index < lengths.length; index++) {
    const end = offset + lengths[index]!;
    if (
      range.start === range.end
        ? range.start >= offset && range.start <= end
        : range.start < end + 1 && range.end > offset
    )
      selected.push(index);
    offset = end + 1;
  }
  return selected;
}
