import type { PreviewMedia } from './media-manifest.ts';

const SVG_NS = 'http://www.w3.org/2000/svg';
const HTML_NS = 'http://www.w3.org/1999/xhtml';
const MAX_TIMEOUT_MS = 2_147_483_647;

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
    clip: PreviewMedia;
    release: () => void;
    retain: (root: HTMLElement) => void;
    element: HTMLMediaElement;
    image: SVGImageElement;
    host: SVGForeignObjectElement;
    shapeId: number;
    play: () => void;
    cancelStart: () => void;
    isTrimEnded: () => boolean;
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
    // PowerPoint's audio "Hide During Show" also hides the playing clip;
    // video becomes visible during playback.
    const hiddenAudio = clip.kind === 'audio' && clip.playback?.hideWhenStopped;
    if (hiddenAudio) host.style.visibility = 'hidden';
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
    let released = false;
    const element = document.createElement(clip.kind);
    const trimStart = Math.max(0, clip.playback?.trim?.startMs ?? 0) / 1000;
    const trimEndRemoved = Math.max(0, clip.playback?.trim?.endMs ?? 0) / 1000;
    const hasTrim = trimStart > 0 || trimEndRemoved > 0;
    const fadeIn = Math.max(0, clip.playback?.fade?.inMs ?? 0) / 1000;
    const fadeOut = Math.max(0, clip.playback?.fade?.outMs ?? 0) / 1000;
    const baseVolume = clip.playback?.volume ?? 0.5;
    let trimEnded = false;
    let metadataReady = false;
    let trimTimer: ReturnType<typeof setTimeout> | undefined;
    let latestTarget: MediaProgress | undefined;
    element.src = clip.src;
    element.preload = 'metadata';
    element.controls = !mirror;
    // Native HTML looping restarts the complete source. PowerPoint loops the
    // trimmed interval, so the boundary is handled by the timeupdate path
    // below whenever a trim is present.
    element.loop = (clip.playback?.loop ?? false) && trimStart === 0 && trimEndRemoved === 0;
    element.volume = baseVolume;
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
    let startTimer: ReturnType<typeof setTimeout> | undefined;
    const cancelStart = () => {
      clearTimeout(startTimer);
      startTimer = undefined;
    };
    const cancelTrimTimer = () => {
      clearTimeout(trimTimer);
      trimTimer = undefined;
    };
    const playableEnd = () => {
      const duration = element.duration;
      if (!Number.isFinite(duration) || duration <= 0) return Number.POSITIVE_INFINITY;
      return Math.max(trimStart, duration - trimEndRemoved);
    };
    const clampTime = (time: number) => Math.min(Math.max(time, trimStart), playableEnd());
    const updateVolume = () => {
      const end = playableEnd();
      const time = Number.isFinite(element.currentTime) ? element.currentTime : trimStart;
      let multiplier = 1;
      if (fadeIn > 0) multiplier = Math.min(multiplier, Math.max(0, (time - trimStart) / fadeIn));
      if (fadeOut > 0 && Number.isFinite(end))
        multiplier = Math.min(multiplier, Math.max(0, (end - time) / fadeOut));
      element.volume = Math.max(0, Math.min(1, baseVolume * multiplier));
    };
    const moveToTrimStart = () => {
      trimEnded = false;
      const target = Number.isFinite(element.duration) ? clampTime(trimStart) : trimStart;
      if (Math.abs(element.currentTime - target) > 0.001) element.currentTime = target;
      updateVolume();
    };
    const finishTrim = () => {
      if (trimEnded || released) return;
      trimEnded = true;
      element.pause();
      const end = playableEnd();
      if (Number.isFinite(end) && Math.abs(element.currentTime - end) > 0.001)
        element.currentTime = end;
      updateVolume();
      // The browser never emits `ended` when we stop at a PowerPoint trim
      // boundary. Keep the same visibility and focus behavior as a native end.
      if (clip.playback?.hideWhenStopped) host.style.visibility = 'hidden';
      if (overlay) {
        const returnFocus = overlay.contains(document.activeElement);
        overlay.hidden = true;
        if (start) start.hidden = false;
        if (returnFocus && !mirror) {
          if (clip.playback?.hideWhenStopped) options.overlayRoot.focus();
          else start?.focus();
        }
      }
      if (clip.playback?.loop) {
        moveToTrimStart();
        host.style.visibility = hiddenAudio ? 'hidden' : '';
        if (overlay) {
          overlay.hidden = false;
          if (start) start.hidden = true;
        }
        void element.play().catch((error) => {
          if (
            disposed ||
            released ||
            (error instanceof DOMException && error.name === 'AbortError')
          )
            return;
          status.textContent = retryLabel;
          status.hidden = false;
        });
        return;
      }
      if (clip.playback?.rewindAfterPlaying) {
        moveToTrimStart();
        // PowerPoint reports the media as completed even though fill=remove
        // has already moved its playhead back to the trimmed start.
        trimEnded = true;
      }
    };
    const keepWithinTrim = () => {
      if (!metadataReady) return;
      if (!hasTrim) {
        updateVolume();
        return;
      }
      if (element.currentTime < trimStart - 0.001) {
        element.currentTime = trimStart;
      } else if (element.currentTime >= playableEnd() - 0.001 && !element.paused) {
        if (mirror) {
          element.pause();
        } else {
          finishTrim();
        }
      }
      updateVolume();
    };
    const scheduleTrimTick = () => {
      cancelTrimTimer();
      if (disposed || released || element.paused || (!hasTrim && fadeIn === 0 && fadeOut === 0))
        return;
      trimTimer = setTimeout(() => {
        keepWithinTrim();
        scheduleTrimTick();
      }, 50);
    };
    const playLocally = async () => {
      cancelStart();
      status.hidden = true;
      if (trimEnded || (metadataReady && element.currentTime < trimStart - 0.001))
        moveToTrimStart();
      keepWithinTrim();
      host.style.visibility = hiddenAudio ? 'hidden' : '';
      if (overlay) {
        overlay.hidden = false;
        if (start) start.hidden = true;
        if (!mirror) element.focus();
      }
      try {
        await element.play();
      } catch (error) {
        if (disposed || released || (error instanceof DOMException && error.name === 'AbortError'))
          return;
        status.textContent =
          error instanceof DOMException && error.name === 'NotAllowedError'
            ? playLabel
            : retryLabel;
        // Keep a way to recover when the browser blocks automatic playback.
        host.style.visibility = '';
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
      if (disposed || released) return;
      status.textContent = retryLabel;
      status.hidden = false;
    });
    element.addEventListener('loadedmetadata', () => {
      metadataReady = true;
      if (!mirror && element.currentTime < trimStart - 0.001) moveToTrimStart();
      updateVolume();
      if (!mirror || !latestTarget) return;
      const targetTime = clampTime(Math.max(0, latestTarget.currentTime));
      if (Math.abs(element.currentTime - targetTime) > 0.3) element.currentTime = targetTime;
      if (!latestTarget.paused && !latestTarget.ended) {
        void element.play().catch(() => {
          if (disposed || released) return;
          status.textContent = retryLabel;
          status.hidden = false;
        });
      }
    });
    element.addEventListener('timeupdate', keepWithinTrim);
    element.addEventListener('seeking', () => {
      if (!metadataReady) return;
      const clamped = clampTime(element.currentTime);
      if (Math.abs(element.currentTime - clamped) > 0.001) element.currentTime = clamped;
      updateVolume();
    });
    for (const name of ['play', 'ended']) {
      element.addEventListener(name, () => {
        if (disposed || released) return;
        if (name === 'play') cancelStart();
        if (name === 'play') scheduleTrimTick();
        if (!mirror && name === 'ended' && hasTrim) {
          finishTrim();
          return;
        }
        if (!mirror && name === 'ended' && clip.playback?.rewindAfterPlaying && !trimEnded) {
          // PowerPoint rewinds after natural completion while keeping the
          // stopped state. Keep the ended/focus handling below unchanged.
          element.pause();
          moveToTrimStart();
        }
        if (!mirror && name === 'ended' && clip.playback?.hideWhenStopped)
          host.style.visibility = 'hidden';
        if (!mirror && name === 'play') host.style.visibility = hiddenAudio ? 'hidden' : '';
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
    element.addEventListener('pause', cancelTrimTimer);
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
      clip,
      release: () => {
        released = true;
        cancelStart();
        cancelTrimTimer();
        if (!mirror && overlay?.contains(document.activeElement)) options.overlayRoot.focus();
        element.pause();
        element.removeAttribute('src');
        element.load();
        host.replaceWith(image);
        overlay?.remove();
        element.remove();
      },
      retain: (root) => {
        cancelStart();
        if (element.parentNode === root) return;
        root.append(element);
        host.replaceWith(image);
        // Moving a media element can pause it in a browser. Resume the same
        // element, preserving its decoder and playback position.
        void playLocally();
      },
      element,
      image,
      host,
      shapeId: clip.shapeId,
      play,
      cancelStart,
      isTrimEnded: () => trimEnded,
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
    if (!mirror && clip.playback?.autoplay) {
      const deadline = performance.now() + (clip.playback.delayMs ?? 0);
      const startWhenDue = () => {
        if (disposed || released) return;
        const remaining = deadline - performance.now();
        // Browser timeouts use signed 32-bit milliseconds; longer OOXML
        // delays must be scheduled in chunks instead of overflowing to zero.
        if (remaining > 0) {
          startTimer = setTimeout(startWhenDue, Math.min(remaining, MAX_TIMEOUT_MS));
        } else void playLocally();
      };
      startWhenDue();
    }
  }
  const entryByShape = new Map(entries.map((entry) => [entry.shapeId, entry]));
  return {
    get progress(): MediaProgress[] {
      return entries.map((entry) => {
        const { element, host, shapeId, overlay } = entry;
        return {
          shapeId,
          currentTime: Number.isFinite(element.currentTime) ? element.currentTime : 0,
          duration: Number.isFinite(element.duration) ? element.duration : 0,
          paused: element.paused,
          ended: element.ended || entry.isTrimEnded(),
          visible: host.style.visibility !== 'hidden',
          fullScreen: Boolean(overlay && !overlay.hidden),
        };
      });
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
        entry.cancelStart();
        entry.element.pause();
      } else if (
        command.action === 'seek' &&
        Number.isFinite(command.time) &&
        (command.time ?? 0) >= 0
      ) {
        const duration = entry.element.duration;
        const trim = entry.clip.playback?.trim;
        const start = Math.max(0, trim?.startMs ?? 0) / 1000;
        const removed = Math.max(0, trim?.endMs ?? 0) / 1000;
        const end = Number.isFinite(duration) ? Math.max(start, duration - removed) : duration;
        entry.element.currentTime = Number.isFinite(end)
          ? Math.min(Math.max(command.time ?? 0, start), end)
          : Math.max(command.time ?? 0, start);
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
        const trim = entry.clip.playback?.trim;
        const start = Math.max(0, trim?.startMs ?? 0) / 1000;
        const removed = Math.max(0, trim?.endMs ?? 0) / 1000;
        const end = Number.isFinite(target.duration)
          ? Math.max(start, target.duration - removed)
          : target.duration;
        const targetTime = Number.isFinite(end)
          ? Math.min(Math.max(start, target.currentTime), end)
          : Math.max(start, target.currentTime);
        if (Math.abs(element.currentTime - targetTime) > 0.3) element.currentTime = targetTime;
        if (target.ended) {
          element.pause();
        }
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
    retainAcrossSlides(distance: number, root: HTMLElement): boolean {
      for (let index = entries.length - 1; index >= 0; index--) {
        const entry = entries[index]!;
        const count = entry.clip.playback?.slideCount ?? 1;
        if (
          !mirror &&
          entry.clip.kind === 'audio' &&
          distance > 0 &&
          distance < count &&
          !entry.element.paused &&
          !entry.element.ended
        ) {
          entry.retain(root);
        } else {
          entry.release();
          entryByShape.delete(entry.shapeId);
          entries.splice(index, 1);
        }
      }
      return entries.length > 0;
    },
    dispose() {
      disposed = true;
      for (const entry of entries) entry.release();
      entries.length = 0;
      entryByShape.clear();
    },
  };
}
