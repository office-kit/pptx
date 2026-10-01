import type { PreviewMedia } from './media-manifest.ts';

const SVG_NS = 'http://www.w3.org/2000/svg';
const HTML_NS = 'http://www.w3.org/1999/xhtml';

export function createMediaPlayer(options: {
  root: ParentNode;
  overlayRoot: HTMLElement;
  clips: readonly PreviewMedia[];
  locale: string;
}) {
  const entries: {
    element: HTMLMediaElement;
    image: SVGImageElement;
    host: SVGForeignObjectElement;
    overlay?: HTMLDivElement;
  }[] = [];
  let disposed = false;
  const retryLabel =
    options.locale === 'ja' ? '再生できませんでした · 再試行' : 'Unable to play media · Retry';
  const playLabel = options.locale === 'ja' ? 'メディアを再生' : 'Play media';
  for (const clip of options.clips) {
    if (clip.kind === 'online') continue;
    const shape = options.root.querySelector(`[data-pptx-shape-id="${clip.shapeId}"]`);
    const image = shape?.querySelector('image');
    if (!(image instanceof SVGImageElement)) continue;
    const host = document.createElementNS(SVG_NS, 'foreignObject');
    for (const name of [
      'x',
      'y',
      'width',
      'height',
      'transform',
      'clip-path',
      'filter',
      'opacity',
    ]) {
      const value = image.getAttribute(name);
      if (value !== null) host.setAttribute(name, value);
    }
    host.dataset.pptxMedia = String(clip.shapeId);
    const box = document.createElementNS(HTML_NS, 'div');
    box.style.cssText = 'position:relative;width:100%;height:100%;';
    const fullScreen = clip.kind === 'video' && clip.playback?.fullScreen;
    const overlay = fullScreen ? document.createElement('div') : undefined;
    let start: HTMLButtonElement | undefined;
    if (overlay) {
      overlay.dataset.pptxMediaFullscreen = String(clip.shapeId);
      overlay.style.cssText = 'position:fixed;inset:0;background:black;z-index:5;';
      overlay.hidden = true;
      options.overlayRoot.append(overlay);
    }
    const element = document.createElement(clip.kind);
    element.src = clip.src;
    element.preload = 'metadata';
    element.controls = true;
    element.loop = clip.playback?.loop ?? false;
    element.volume = clip.playback?.volume ?? 0.5;
    element.muted = clip.playback?.muted ?? false;
    element.setAttribute('aria-label', clip.kind === 'video' ? 'Video' : 'Audio');
    element.style.cssText = 'display:block;width:100%;height:100%;object-fit:fill;';
    if (overlay) element.style.objectFit = 'contain';
    if (element instanceof HTMLVideoElement) {
      element.playsInline = true;
      element.poster = image.href.baseVal;
    } else {
      box.style.background = `center / contain no-repeat url("${image.href.baseVal}")`;
    }
    const status = document.createElement('button');
    status.type = 'button';
    status.hidden = true;
    status.style.cssText =
      'position:absolute;inset:0;margin:auto;max-width:100%;height:fit-content;';
    const play = async () => {
      status.hidden = true;
      host.style.visibility = '';
      if (overlay) {
        overlay.hidden = false;
        element.focus();
      }
      try {
        await element.play();
      } catch (error) {
        if (disposed || (error instanceof DOMException && error.name === 'AbortError')) return;
        status.textContent =
          error instanceof DOMException && error.name === 'NotAllowedError'
            ? playLabel
            : retryLabel;
        status.hidden = false;
      }
    };
    status.onclick = () => {
      if (element.error) element.load();
      void play();
    };
    for (const name of ['click', 'dblclick', 'pointerdown', 'keydown']) {
      const stopNavigation = (event: Event) => {
        if (!(event instanceof KeyboardEvent) || event.key !== 'Escape') event.stopPropagation();
      };
      box.addEventListener(name, stopNavigation);
      overlay?.addEventListener(name, stopNavigation);
    }
    element.addEventListener('error', () => {
      if (disposed) return;
      status.textContent = retryLabel;
      status.hidden = false;
    });
    for (const name of ['play', 'ended']) {
      element.addEventListener(name, () => {
        if (disposed) return;
        if (name === 'ended' && clip.playback?.hideWhenStopped) host.style.visibility = 'hidden';
        if (name === 'play') host.style.visibility = '';
        if (overlay) {
          const returnFocus = name === 'ended' && overlay.contains(document.activeElement);
          overlay.hidden = name === 'ended';
          if (returnFocus) {
            if (clip.playback?.hideWhenStopped) options.overlayRoot.focus();
            else start?.focus();
          }
        }
      });
    }
    if (overlay) {
      box.style.background = `center / contain no-repeat url("${image.href.baseVal}")`;
      start = document.createElement('button');
      start.type = 'button';
      start.setAttribute('aria-label', playLabel);
      start.title = playLabel;
      start.style.cssText =
        'position:absolute;inset:0;width:100%;height:100%;padding:0;border:0;background:transparent;cursor:pointer;';
      start.onclick = () => {
        void play();
      };
      box.append(start);
      overlay.append(element, status);
    } else box.append(element, status);
    host.append(box);
    image.replaceWith(host);
    entries.push({ element, image, host, ...(overlay ? { overlay } : {}) });
    if (clip.playback?.autoplay) void play();
  }
  return {
    dispose() {
      disposed = true;
      for (const { element, image, host, overlay } of entries) {
        if (overlay?.contains(document.activeElement)) options.overlayRoot.focus();
        element.pause();
        element.removeAttribute('src');
        element.load();
        host.replaceWith(image);
        overlay?.remove();
      }
    },
  };
}
