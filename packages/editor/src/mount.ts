import { mount, unmount, type Component } from 'svelte';
import contextualTabs from './ribbon/contextual.css?inline';
import tokens from './ui/tokens.css?inline';

// The editor's global stylesheets; component styles are injected by Svelte.
const STYLES = tokens + contextualTabs;

// The host element is a (never registered) custom element so that page rules for
// `div` and the like do not match it, and inline `all: initial` stops the page's
// inherited properties (font, line-height, letter-spacing …) at the boundary.
const HOST_TAG = 'office-kit-pptx-editor';
const HOST_STYLE = 'all: initial; display: block; position: relative; width: 100%; height: 100%;';

/**
 * Mounts `component` in a shadow root appended to `target`: the editor's styles
 * stay inside it and the page's stay out. The returned function removes the host
 * element again, leaving the target as it was so it can be mounted again.
 */
export function mountInShadowRoot<Props extends Record<string, unknown>>(
  target: HTMLElement,
  component: Component<Props>,
  props: Props,
): () => void {
  const host = target.ownerDocument.createElement(HOST_TAG);
  host.style.cssText = HOST_STYLE;
  const shadow = host.attachShadow({ mode: 'open' });
  const style = target.ownerDocument.createElement('style');
  style.textContent = STYLES;
  shadow.append(style);
  target.append(host);
  const app = mount(component, { target: shadow, props });
  return () => {
    void unmount(app);
    host.remove();
  };
}

/**
 * Mounts `component` directly in the page, which then shares the editor's
 * styles. The returned function removes it and its styles again.
 */
export function mountInPage<Props extends Record<string, unknown>>(
  target: HTMLElement,
  component: Component<Props>,
  props: Props,
): () => void {
  const style = target.ownerDocument.createElement('style');
  style.textContent = STYLES;
  target.ownerDocument.head.append(style);
  const app = mount(component, { target, props });
  return () => {
    void unmount(app);
    style.remove();
  };
}
