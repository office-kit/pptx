import type { EditorShape } from './editor.ts';

/** Crop edges follow the picture's local axes; source crop follows its unflipped axes. */
export function dragPictureCrop(
  shape: EditorShape,
  handle: string,
  dx: number,
  dy: number,
  aspectRatio?: number,
) {
  const original = shape.bounds!;
  const angle = (shape.rotation * Math.PI) / 180,
    c = Math.cos(angle),
    s = Math.sin(angle);
  const x = dx * c + dy * s,
    y = -dx * s + dy * c;
  const crop = { left: 0, right: 0, top: 0, bottom: 0, ...shape.imageCrop };
  const sourceW = original.w / (1 - crop.left - crop.right);
  const sourceH = original.h / (1 - crop.top - crop.bottom);
  const horizontal = shape.flip?.horizontal ? -1 : 1;
  const vertical = shape.flip?.vertical ? -1 : 1;
  if (handle.startsWith('source-')) {
    const corner = handle.slice(7);
    const left = -(horizontal === 1 ? crop.left : crop.right) * sourceW;
    const top = -(vertical === 1 ? crop.top : crop.bottom) * sourceH;
    const widthChange = (corner.includes('w') ? -x : x) / sourceW;
    const heightChange = (corner.includes('n') ? -y : y) / sourceH;
    const scale = Math.max(
      12700 / Math.min(sourceW, sourceH),
      1 + (Math.abs(widthChange) >= Math.abs(heightChange) ? widthChange : heightChange),
    );
    const width = sourceW * scale,
      height = sourceH * scale;
    const nextLeft = corner.includes('w') ? left + sourceW - width : left;
    const nextTop = corner.includes('n') ? top + sourceH - height : top;
    crop[horizontal === 1 ? 'left' : 'right'] = -nextLeft / width;
    crop[horizontal === 1 ? 'right' : 'left'] = 1 - (original.w - nextLeft) / width;
    crop[vertical === 1 ? 'top' : 'bottom'] = -nextTop / height;
    crop[vertical === 1 ? 'bottom' : 'top'] = 1 - (original.h - nextTop) / height;
    return { bounds: original, imageCrop: crop };
  }
  if (!handle) {
    crop.left -= (x / sourceW) * horizontal;
    crop.right += (x / sourceW) * horizontal;
    crop.top -= (y / sourceH) * vertical;
    crop.bottom += (y / sourceH) * vertical;
    return { bounds: original, imageCrop: crop };
  }
  let left = 0,
    right = original.w,
    top = 0,
    bottom = original.h;
  const minimum = Math.min(12700, original.w, original.h);
  if (handle.includes('w')) left = Math.min(x, right - minimum);
  if (handle.includes('e')) right = Math.max(original.w + x, left + minimum);
  if (handle.includes('n')) top = Math.min(y, bottom - minimum);
  if (handle.includes('s')) bottom = Math.max(original.h + y, top + minimum);
  if (aspectRatio !== undefined) {
    if (!Number.isFinite(aspectRatio) || aspectRatio <= 0)
      throw new Error('Invalid crop aspect ratio.');
    const horizontalHandle = handle.includes('e') || handle.includes('w');
    const verticalHandle = handle.includes('n') || handle.includes('s');
    const proposedW = right - left,
      proposedH = bottom - top;
    const useWidth =
      horizontalHandle &&
      (!verticalHandle ||
        Math.abs(proposedW / original.w - 1) >= Math.abs(proposedH / original.h - 1));
    const width = Math.max(
      minimum,
      minimum * aspectRatio,
      useWidth ? proposedW : proposedH * aspectRatio,
    );
    const height = width / aspectRatio;
    if (handle.includes('w')) left = right - width;
    else if (handle.includes('e')) right = left + width;
    else {
      left = (original.w - width) / 2;
      right = left + width;
    }
    if (handle.includes('n')) top = bottom - height;
    else if (handle.includes('s')) bottom = top + height;
    else {
      top = (original.h - height) / 2;
      bottom = top + height;
    }
  }
  crop[horizontal === 1 ? 'left' : 'right'] += left / sourceW;
  crop[horizontal === 1 ? 'right' : 'left'] += (original.w - right) / sourceW;
  crop[vertical === 1 ? 'top' : 'bottom'] += top / sourceH;
  crop[vertical === 1 ? 'bottom' : 'top'] += (original.h - bottom) / sourceH;
  const w = right - left,
    h = bottom - top;
  const cx = (left + right - original.w) / 2,
    cy = (top + bottom - original.h) / 2;
  return {
    bounds: {
      x: original.x + original.w / 2 + cx * c - cy * s - w / 2,
      y: original.y + original.h / 2 + cx * s + cy * c - h / 2,
      w,
      h,
    },
    imageCrop: crop,
  };
}
