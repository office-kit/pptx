import type { PreviewMedia } from './media-manifest.ts';

const SVG_NS = 'http://www.w3.org/2000/svg';
const HTML_NS = 'http://www.w3.org/1999/xhtml';

export type MediaCommand = {
  shapeId: number;
  action: 'play' | 'pause' | 'seek';
  time?: number;
};

export type MediaProgress = {
  shapeId: number;
  currentTime: number;
  duration: number;
  paused: boolean;
  ended: boolean;
  visible: boolean;
  fullScreen: boolean;
};

export function createMediaPlayer(options: {
  root: ParentNode;
  overlayRoot: HTMLElement;
  clips: readonly PreviewMedia[];
  locale: string;
  /** Presence makes this instance a presenter mirror. Commands are forwarded
   * to the audience instead of changing the mirror independently. */
  onCommand?: (command: MediaCommand) => void;
}) {
  const mirror = options.onCommand !== undefined;
  const entries: {
    element: HTMLMediaElement;
    image: SVGImageElement;
    host: SVGForeignObjectElement;
    shapeId: number;
    play: () => void;
    status: HTMLButtonElement;
    setTarget: (target: MediaProgress) => void;
    start?: HTMLButtonElement;
    controls?: {
      toggle: HTMLButtonElement;
      position: HTMLInputElement;
    };
    overlay?: HTMLDivElement;
  }[] = [];
  let disposed = false;
  const retryLabel =
    options.locale === 'ja' ? '再生できませんでした · 再試行' : 'Unable to play media · Retry';
  const playLabel = options.locale === 'ja' ? 'メディアを再生' : 'Play media';
  const pauseLabel = options.locale === 'ja' ? 'メディアを一時停止' : 'Pause media';
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
      overlay.style.cssText = `position:${mirror ? 'absolute' : 'fixed'};inset:0;background:black;z-index:5;`;
      overlay.hidden = true;
      options.overlayRoot.append(overlay);
    }
    const element = document.createElement(clip.kind);
    let latestTarget: MediaProgress | undefined;
    element.src = clip.src;
    element.preload = 'metadata';
    element.controls = !mirror;
    element.loop = clip.playback?.loop ?? false;
    element.volume = clip.playback?.volume ?? 0.5;
    // A presenter mirror decodes the same clip silently. Its controls send a
    // command to the audience, while sync() is the only source of playback
    // state, so it must never autoplay with sound.
    element.muted = mirror || (clip.playback?.muted ?? false);
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
    let controls: { toggle: HTMLButtonElement; position: HTMLInputElement } | undefined;
    const playLocally = async () => {
      status.hidden = true;
      host.style.visibility = '';
      if (overlay) {
        overlay.hidden = false;
        if (start) start.hidden = true;
        if (!mirror) element.focus();
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
    const play = () => {
      if (mirror) {
        options.onCommand?.({ shapeId: clip.shapeId, action: 'play' });
        return;
      }
      void playLocally();
    };
    if (mirror) {
      const playButton = document.createElement('button');
      playButton.type = 'button';
      playButton.textContent = playLabel;
      playButton.setAttribute('aria-label', playLabel);
      const position = document.createElement('input');
      position.type = 'range';
      position.min = '0';
      position.max = '100';
      position.step = '0.1';
      position.value = '0';
      position.setAttribute(
        'aria-label',
        options.locale === 'ja' ? 'メディア位置' : 'Media position',
      );
      const controlsRoot = document.createElement('div');
      controlsRoot.style.cssText =
        'position:absolute;left:4px;right:4px;bottom:4px;z-index:2;display:flex;align-items:center;gap:4px;padding:3px;background:rgb(0 0 0 / 65%);';
      position.style.cssText = 'min-width:0;flex:1;';
      playButton.onclick = () => {
        const target = latestTarget;
        if (target ? !target.paused && !target.ended : !element.paused) {
          options.onCommand?.({ shapeId: clip.shapeId, action: 'pause' });
        } else {
          play();
        }
      };
      position.oninput = () => {
        const duration = latestTarget?.duration || element.duration;
        const percent = Number(position.value);
        if (Number.isFinite(duration) && duration > 0 && Number.isFinite(percent)) {
          options.onCommand?.({
            shapeId: clip.shapeId,
            action: 'seek',
            time: (duration * percent) / 100,
          });
        }
      };
      controlsRoot.append(playButton, position);
      controls = { toggle: playButton, position };
      (overlay ?? box).append(controlsRoot);
    }
    status.onclick = () => {
      if (mirror) status.hidden = true;
      if (element.error) element.load();
      play();
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
    element.addEventListener('loadedmetadata', () => {
      if (!mirror || !latestTarget) return;
      const targetTime = Math.max(0, latestTarget.currentTime);
      if (Math.abs(element.currentTime - targetTime) > 0.3) element.currentTime = targetTime;
      if (!latestTarget.paused && !latestTarget.ended) {
        void element.play().catch(() => {
          if (disposed) return;
          status.textContent = retryLabel;
          status.hidden = false;
        });
      }
    });
    for (const name of ['play', 'ended']) {
      element.addEventListener(name, () => {
        if (disposed) return;
        if (!mirror && name === 'ended' && clip.playback?.hideWhenStopped)
          host.style.visibility = 'hidden';
        if (!mirror && name === 'play') host.style.visibility = '';
        if (overlay) {
          const returnFocus = name === 'ended' && overlay.contains(document.activeElement);
          if (!mirror) overlay.hidden = name === 'ended';
          if (start && name === 'ended') start.hidden = false;
          if (returnFocus && !mirror) {
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
        play();
      };
      box.append(start);
      overlay.append(element, status);
    } else box.append(element, status);
    host.append(box);
    image.replaceWith(host);
    entries.push({
      element,
      image,
      host,
      shapeId: clip.shapeId,
      play,
      status,
      ...(start ? { start } : {}),
      ...(controls ? { controls } : {}),
      setTarget: (target) => {
        latestTarget = target;
      },
      ...(overlay ? { overlay } : {}),
    });
    // A mirror starts only after the first audience snapshot. Calling play
    // here would race that snapshot and can also violate autoplay policy.
    if (!mirror && clip.playback?.autoplay) void playLocally();
  }
  const entryByShape = new Map(entries.map((entry) => [entry.shapeId, entry]));
  return {
    get progress(): MediaProgress[] {
      return entries.map(({ element, host, shapeId, overlay }) => ({
        shapeId,
        currentTime: Number.isFinite(element.currentTime) ? element.currentTime : 0,
        duration: Number.isFinite(element.duration) ? element.duration : 0,
        paused: element.paused,
        ended: element.ended,
        visible: host.style.visibility !== 'hidden',
        fullScreen: Boolean(overlay && !overlay.hidden),
      }));
    },
    command(command: MediaCommand) {
      const entry = entryByShape.get(command.shapeId);
      if (!entry) return;
      if (mirror) {
        if (
          command.action === 'seek' &&
          (!Number.isFinite(command.time) || (command.time ?? 0) < 0)
        )
          return;
        options.onCommand?.(command);
        return;
      }
      if (command.action === 'play') {
        entry.play();
      } else if (command.action === 'pause') {
        entry.element.pause();
      } else if (
        command.action === 'seek' &&
        Number.isFinite(command.time) &&
        (command.time ?? 0) >= 0
      ) {
        const duration = entry.element.duration;
        entry.element.currentTime = Number.isFinite(duration)
          ? Math.min(command.time ?? 0, duration)
          : (command.time ?? 0);
      }
    },
    sync(progress: readonly MediaProgress[]) {
      if (!mirror) return;
      for (const target of progress) {
        const entry = entryByShape.get(target.shapeId);
        if (!entry) continue;
        const { element, host, overlay } = entry;
        entry.setTarget(target);
        if (entry.start) entry.start.hidden = target.fullScreen;
        if (entry.controls) {
          const playing = !target.paused && !target.ended;
          entry.controls.toggle.textContent = playing ? pauseLabel : playLabel;
          entry.controls.toggle.setAttribute('aria-label', playing ? pauseLabel : playLabel);
          const duration = target.duration;
          entry.controls.position.value =
            duration > 0
              ? String(Math.max(0, Math.min(100, (target.currentTime / duration) * 100)))
              : '0';
        }
        const targetTime = Math.max(0, target.currentTime);
        if (Math.abs(element.currentTime - targetTime) > 0.3) element.currentTime = targetTime;
        host.style.visibility = target.visible ? '' : 'hidden';
        if (overlay) overlay.hidden = !target.fullScreen;
        if (target.paused || target.ended) {
          element.pause();
        } else if (element.paused && entry.status.hidden) {
          void element.play().catch((error) => {
            if (disposed) return;
            entry.status.textContent =
              error instanceof DOMException && error.name === 'NotAllowedError'
                ? playLabel
                : retryLabel;
            entry.status.hidden = false;
          });
        }
      }
    },
    dispose() {
      disposed = true;
      for (const { element, image, host, overlay } of entries) {
        if (!mirror && overlay?.contains(document.activeElement)) options.overlayRoot.focus();
        element.pause();
        element.removeAttribute('src');
        element.load();
        host.replaceWith(image);
        overlay?.remove();
      }
    },
  };
}
