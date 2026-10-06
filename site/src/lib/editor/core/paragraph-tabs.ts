import type { ParagraphTabStop } from '@office-kit/pptx';

export type TabStopEdit =
  | { kind: 'set'; stop: ParagraphTabStop }
  | { kind: 'clear'; positionEmu: number }
  | { kind: 'clearAll' };

/** Apply only the requested tab changes, preserving other stops in mixed selections. */
export function editTabStops(
  stops: readonly ParagraphTabStop[],
  edits: readonly TabStopEdit[],
): ParagraphTabStop[] {
  const result = new Map(stops.map((stop) => [stop.positionEmu, stop]));
  for (const edit of edits) {
    if (edit.kind === 'clearAll') result.clear();
    else if (edit.kind === 'clear') result.delete(edit.positionEmu);
    else result.set(edit.stop.positionEmu, edit.stop);
  }
  return [...result.values()].sort((a, b) => a.positionEmu - b.positionEmu);
}
