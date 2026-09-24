/** Physical codes survive the alternate characters produced by Option on macOS. */
export function arrangeShortcut(
  event: Pick<KeyboardEvent, 'metaKey' | 'ctrlKey' | 'altKey' | 'shiftKey' | 'code'>,
): string | null {
  if (!event.metaKey && !event.ctrlKey) return null;
  if (event.code === 'KeyJ' && event.altKey && !event.shiftKey) return 'regroup';
  if (event.code === 'KeyG') return event.shiftKey ? 'ungroupShapes' : 'groupShapes';
  if (!event.shiftKey) return null;
  if (event.code === 'KeyF') return event.altKey ? 'bringShapeForward' : 'bringShapeToFront';
  if (event.code === 'KeyB') return event.altKey ? 'sendShapeBackward' : 'sendShapeToBack';
  return null;
}
