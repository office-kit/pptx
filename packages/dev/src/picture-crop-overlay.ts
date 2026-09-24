import type { EditorShape } from './editor.ts';

/** Show the uncropped source only outside the active crop rectangle. */
export function renderPictureCropOverlay(
  hit: HTMLElement,
  shape: EditorShape,
  source: SVGImageElement | null,
) {
  hit.querySelector('.picture-crop-source')?.remove();
  hit.querySelector('.picture-source-frame')?.remove();
  if (!source || !shape.bounds) return;
  const { w, h } = shape.bounds;
  const crop = { left: 0, right: 0, top: 0, bottom: 0, ...shape.imageCrop };
  const width = w / (1 - crop.left - crop.right),
    height = h / (1 - crop.top - crop.bottom);
  const x = -crop.left * width,
    y = -crop.top * height;
  if (![x, y, width, height].every(Number.isFinite) || width <= 0 || height <= 0) return;
  const ns = 'http://www.w3.org/2000/svg';
  const make = (tag: string, attrs: Record<string, string | number>) => {
    const node = document.createElementNS(ns, tag);
    for (const [key, value] of Object.entries(attrs)) node.setAttribute(key, String(value));
    return node;
  };
  const svg = make('svg', { viewBox: `0 0 ${w} ${h}`, 'aria-hidden': 'true' });
  svg.classList.add('picture-crop-source');
  svg.style.cssText =
    'position:absolute;inset:0;width:100%;height:100%;overflow:visible;pointer-events:none';
  const flipX = shape.flip?.horizontal,
    flipY = shape.flip?.vertical;
  const visualX = flipX ? w - x - width : x,
    visualY = flipY ? h - y - height : y;
  const maskId = `picture-crop-mask-${shape.id}`;
  const mask = make('mask', {
    id: maskId,
    maskUnits: 'userSpaceOnUse',
    maskContentUnits: 'userSpaceOnUse',
    x: visualX,
    y: visualY,
    width,
    height,
    'mask-type': 'luminance',
  });
  mask.append(
    make('rect', { x: visualX, y: visualY, width, height, fill: 'white' }),
    make('rect', { x: 0, y: 0, width: w, height: h, fill: 'black' }),
  );
  const defs = make('defs', {});
  defs.append(mask);
  const outside = make('g', { mask: `url(#${maskId})` });
  const transformed = make('g', {
    transform: `translate(${flipX ? w : 0} ${flipY ? h : 0}) scale(${flipX ? -1 : 1} ${flipY ? -1 : 1})`,
  });
  transformed.style.filter = 'brightness(0.45)';
  const image = source.cloneNode(true) as SVGImageElement;
  image.removeAttribute('id');
  for (const [key, value] of Object.entries({ x, y, width, height }))
    image.setAttribute(key, String(value));
  transformed.append(image);
  outside.append(transformed);
  svg.append(defs, outside);
  const frame = document.createElement('div');
  frame.className = 'picture-source-frame';
  Object.assign(frame.style, {
    position: 'absolute',
    pointerEvents: 'none',
    left: `${(visualX / w) * 100}%`,
    top: `${(visualY / h) * 100}%`,
    width: `${(width / w) * 100}%`,
    height: `${(height / h) * 100}%`,
  });
  for (const corner of ['nw', 'ne', 'sw', 'se']) {
    const handle = document.createElement('span');
    handle.className = 'picture-source-handle';
    handle.dataset.handle = `source-${corner}`;
    handle.setAttribute('aria-label', `Resize picture ${corner}`);
    Object.assign(handle.style, {
      left: corner.includes('w') ? '0' : '100%',
      top: corner.includes('n') ? '0' : '100%',
      cursor: ['nw', 'se'].includes(corner) ? 'nwse-resize' : 'nesw-resize',
    });
    frame.append(handle);
  }
  hit.prepend(frame);
  // Handles must paint above the dimmed source.
  hit.prepend(svg);
}
