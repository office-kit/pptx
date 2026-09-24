import type { EditCommand } from './editor.ts';

/** A disposable server-rendered preview; the endpoint does not persist preview edits. */
export function previewPictureMenu(options: {
  command: EditCommand;
  revision: number;
  current(): boolean;
  shadow: ShadowRoot;
  layer: HTMLElement;
}) {
  const controller = new AbortController();
  let overlay: HTMLDivElement | undefined;
  let original: SVGSVGElement | undefined;
  let visibility = '';
  const timer = setTimeout(async () => {
    try {
      const response = await fetch('/edit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        signal: controller.signal,
        body: JSON.stringify({
          revision: options.revision,
          command: options.command,
          preview: true,
          previewRender: true,
        }),
      });
      if (!response.ok) return;
      const result = await response.json();
      if (controller.signal.aborted || !options.current() || !result.svg) return;
      original = [...options.shadow.children].find(
        (node): node is SVGSVGElement => node instanceof SVGSVGElement,
      );
      if (!original) return;
      overlay = document.createElement('div');
      overlay.className = 'picture-menu-preview';
      overlay.style.cssText = 'position:absolute;inset:0;pointer-events:none';
      overlay.innerHTML = result.svg;
      const ids = new Map(
        [...overlay.querySelectorAll('[id]')].map((node) => [
          node.id,
          'picture-menu-preview-' + node.id,
        ]),
      );
      for (const node of overlay.querySelectorAll('*')) {
        for (const attribute of node.attributes) {
          let value = attribute.value.replace(/url\(#([^)]*)\)/g, (whole, id) =>
            ids.has(id) ? `url(#${ids.get(id)})` : whole,
          );
          if (attribute.localName === 'href' && value.startsWith('#') && ids.has(value.slice(1)))
            value = '#' + ids.get(value.slice(1));
          if (attribute.name === 'id') value = ids.get(value) ?? value;
          if (value !== attribute.value)
            node.setAttributeNS(attribute.namespaceURI, attribute.name, value);
        }
      }
      const svg = overlay.querySelector('svg');
      if (svg) {
        svg.style.width = '100%';
        svg.style.height = '100%';
      }
      visibility = original.style.visibility;
      original.style.visibility = 'hidden';
      options.layer.prepend(overlay);
    } catch {
      // Closing or moving between candidates aborts obsolete renders.
    }
  }, 80);
  let disposed = false;
  return () => {
    if (disposed) return;
    disposed = true;
    clearTimeout(timer);
    controller.abort();
    overlay?.remove();
    if (original) original.style.visibility = visibility;
  };
}
