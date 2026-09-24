export interface ConnectorFrame {
  bounds: { x: number; y: number; w: number; h: number } | null;
  rotation: number;
  flip: { horizontal: boolean; vertical: boolean } | null;
}
export function shapePoint(frame: ConnectorFrame, point: { x: number; y: number }) {
  const b = frame.bounds!;
  const x = (frame.flip?.horizontal ? b.w - point.x : point.x) - b.w / 2;
  const y = (frame.flip?.vertical ? b.h - point.y : point.y) - b.h / 2;
  const a = (frame.rotation * Math.PI) / 180;
  return {
    x: b.x + b.w / 2 + x * Math.cos(a) - y * Math.sin(a),
    y: b.y + b.h / 2 + x * Math.sin(a) + y * Math.cos(a),
  };
}
export function connectorEndpoints(frame: ConnectorFrame) {
  return {
    start: shapePoint(frame, { x: 0, y: 0 }),
    end: shapePoint(frame, { x: frame.bounds!.w, y: frame.bounds!.h }),
  };
}
/** A rotated connector's frame, preserving endpoint identity even when the ends cross. */
export function connectorFrame(
  start: { x: number; y: number },
  end: { x: number; y: number },
  rotation: number,
) {
  const a = (rotation * Math.PI) / 180;
  const dx = end.x - start.x,
    dy = end.y - start.y;
  const w = Math.round(dx * Math.cos(a) + dy * Math.sin(a));
  const h = Math.round(-dx * Math.sin(a) + dy * Math.cos(a));
  return {
    bounds: {
      x: Math.round((start.x + end.x - Math.abs(w)) / 2),
      y: Math.round((start.y + end.y - Math.abs(h)) / 2),
      w: Math.abs(w),
      h: Math.abs(h),
    },
    flip: { horizontal: w < 0, vertical: h < 0 },
  };
}
