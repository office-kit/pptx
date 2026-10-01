import type { MediaPlayback } from '@office-kit/pptx';
import type { EditorController } from './controller.svelte.ts';

export type MediaPreviewSnapshot = {
  shapeId: number | null;
  currentTime: number;
  duration: number;
  playing: boolean;
  bookmarkIndex: number | null;
  error: string;
};
export type MediaPreviewController = {
  state: MediaPreviewSnapshot;
  attach(shapeId: number, element: HTMLMediaElement): () => void;
  reset(): void;
  command(command: 'play' | 'pause' | 'seek', shapeId: number, time?: number): void;
  selectBookmark(index: number | null, timeMs?: number): void;
  bookmarkAtCurrent(playback: MediaPlayback | null): number | null;
};

const bookmarkTimeToleranceMs = 0.001;
const controllers = new WeakMap<object, MediaPreviewController>();

export function getMediaPreview(editor: EditorController): MediaPreviewController {
  const existing = controllers.get(editor);
  if (existing) return existing;
  let element: HTMLMediaElement | null = null;
  let detach: (() => void) | null = null;
  let selectedTimeMs: number | null = null;
  const state = $state<MediaPreviewSnapshot>({
    shapeId: null,
    currentTime: 0,
    duration: 0,
    playing: false,
    bookmarkIndex: null,
    error: '',
  });
  const update = () => {
    if (!element) return;
    state.currentTime = Number.isFinite(element.currentTime) ? element.currentTime : 0;
    state.duration = Number.isFinite(element.duration) ? element.duration : 0;
    state.playing = !element.paused;
    if (
      selectedTimeMs !== null &&
      Math.abs(state.currentTime * 1000 - selectedTimeMs) > bookmarkTimeToleranceMs
    ) {
      state.bookmarkIndex = null;
      selectedTimeMs = null;
    }
  };
  const controller: MediaPreviewController = {
    state,
    attach(shapeId, next) {
      detach?.();
      element = next;
      selectedTimeMs = null;
      state.shapeId = shapeId;
      state.bookmarkIndex = null;
      state.error = '';
      const events = [
        'timeupdate',
        'durationchange',
        'loadedmetadata',
        'play',
        'pause',
        'ended',
      ] as const;
      events.forEach((name) => next.addEventListener(name, update));
      update();
      const cleanup = () => {
        events.forEach((name) => next.removeEventListener(name, update));
        next.pause();
        if (element !== next) return;
        element = null;
        state.shapeId = null;
        state.playing = false;
        state.bookmarkIndex = null;
        state.error = '';
        selectedTimeMs = null;
        if (detach === cleanup) detach = null;
      };
      detach = cleanup;
      return detach;
    },
    reset() {
      detach?.();
    },
    command(command, shapeId, time) {
      if (!element || state.shapeId !== shapeId) return;
      state.bookmarkIndex = null;
      selectedTimeMs = null;
      state.error = '';
      const target = element;
      const attachment = detach;
      if (command === 'seek' && time !== undefined) element.currentTime = Math.max(0, time);
      else if (command === 'play')
        void element.play().catch((error) => {
          if (element !== target || state.shapeId !== shapeId || detach !== attachment) return;
          state.playing = false;
          state.error = error instanceof Error ? error.message : String(error);
        });
      else if (command === 'pause') element.pause();
      update();
    },
    selectBookmark(index, timeMs) {
      state.bookmarkIndex = index;
      selectedTimeMs = index === null ? null : (timeMs ?? state.currentTime * 1000);
    },
    bookmarkAtCurrent(playback) {
      if (!playback?.bookmarks?.length) return null;
      const current = state.currentTime * 1000;
      const index = playback.bookmarks.findIndex(
        (bookmark) => Math.abs(bookmark.timeMs - current) <= bookmarkTimeToleranceMs,
      );
      return index < 0 ? null : index;
    },
  };
  controllers.set(editor, controller);
  return controller;
}
