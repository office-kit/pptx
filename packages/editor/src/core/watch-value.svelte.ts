import { untrack } from 'svelte';

/**
 * Calls `listener` whenever the value `read` returns changes. Returns a
 * function that stops watching.
 */
export function watchValue<T>(read: () => T, listener: (value: T) => void): () => void {
  let last = untrack(read);
  return $effect.root(() => {
    $effect(() => {
      const value = read();
      if (value === last) return;
      last = value;
      untrack(() => listener(value));
    });
  });
}
