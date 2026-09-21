const MAX_PANES = 4;
type Layout =
  | { id: string }
  | { direction: 'horizontal' | 'vertical'; ratio: number; first: Layout; second: Layout };

export function mountWorkspace() {
  const host = document.getElementById('agent-workspace')!;
  const key = 'office-kit-agent-layout';
  const panes = new Map<string, HTMLElement>();
  const frames = new Map<string, HTMLIFrameElement>();
  let layout: Layout = { id: crypto.randomUUID() };
  let serial = 0;
  const valid = (value: unknown, ids: Set<string>): value is Layout => {
    if (!value || typeof value !== 'object') return false;
    if ('id' in value) {
      if (typeof value.id !== 'string' || !/^[\w-]{1,64}$/.test(value.id) || ids.has(value.id))
        return false;
      ids.add(value.id);
      return ids.size <= MAX_PANES;
    }
    return (
      'direction' in value &&
      ['horizontal', 'vertical'].includes(String(value.direction)) &&
      'ratio' in value &&
      typeof value.ratio === 'number' &&
      value.ratio >= 0.15 &&
      value.ratio <= 0.85 &&
      'first' in value &&
      'second' in value &&
      valid(value.first, ids) &&
      valid(value.second, ids)
    );
  };
  try {
    const saved: unknown = JSON.parse(sessionStorage.getItem(key) || 'null');
    if (valid(saved, new Set())) layout = saved;
  } catch {
    /* Storage may be disabled, or contain an older layout. */
  }
  function save() {
    try {
      sessionStorage.setItem(key, JSON.stringify(layout));
    } catch {
      /* Layout remains usable without storage. */
    }
  }
  function broadcastFocus() {
    const context = document.getElementById('chat-context')!;
    for (const frame of frames.values())
      frame.contentWindow?.postMessage(
        {
          type: 'focus',
          focus: JSON.parse(context.dataset.focus || '{"slide":null,"revision":0}'),
          label: context.textContent,
        },
        location.origin,
      );
  }
  window.addEventListener('agent-focus', broadcastFocus);
  window.addEventListener('agent-chat', () => {
    for (const frame of frames.values())
      frame.contentWindow?.postMessage({ type: 'chat' }, location.origin);
  });
  window.addEventListener('message', (event) => {
    if (
      event.origin === location.origin &&
      [...frames.values()].some((frame) => frame.contentWindow === event.source) &&
      event.data?.type === 'agent-ready'
    )
      broadcastFocus();
  });
  function replace(node: Layout, id: string, next: Layout | null): Layout | null {
    if ('id' in node) return node.id === id ? next : node;
    const first = replace(node.first, id, next),
      second = replace(node.second, id, next);
    if (!first) return second;
    if (!second) return first;
    return { ...node, first, second };
  }
  function button(label: string, action: () => void) {
    const button = document.createElement('button');
    button.setAttribute('aria-label', label);
    button.title = label;
    const path =
      label === 'Split right'
        ? '<rect x="3" y="3" width="18" height="18" rx="3"/><path d="M12 3v18"/>'
        : label === 'Split down'
          ? '<rect x="3" y="3" width="18" height="18" rx="3"/><path d="M3 12h18"/>'
          : '<path d="m6 6 12 12M18 6 6 18"/>';
    button.innerHTML =
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" aria-hidden="true">' +
      path +
      '</svg>';
    button.onclick = action;
    return button;
  }
  function pane(id: string) {
    const existing = panes.get(id);
    if (existing) return existing;
    const section = document.createElement('section'),
      tools = document.createElement('div'),
      title = document.createElement('span');
    section.className = 'agent-pane';
    tools.className = 'agent-tools';
    const label = 'Agent ' + ++serial;
    section.setAttribute('aria-label', label);
    title.textContent = label;
    tools.append(title);
    for (const direction of ['horizontal', 'vertical'] as const) {
      const split = button(direction === 'horizontal' ? 'Split right' : 'Split down', () => {
        if (panes.size >= MAX_PANES) return;
        layout = replace(layout, id, {
          direction,
          ratio: 0.5,
          first: { id },
          second: { id: crypto.randomUUID() },
        })!;
        render();
      });
      split.dataset.split = direction;
      tools.append(split);
    }
    const close = button('Close', () => {
      void closePane();
    });
    close.dataset.close = '';
    async function closePane() {
      close.disabled = true;
      try {
        const response = await fetch('/agents/' + id + '/close', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: '{}',
        });
        if (!response.ok) throw new Error('Could not close agent. Try again.');
        layout = replace(layout, id, null) || { id: crypto.randomUUID() };
        frames.delete(id);
        panes.delete(id);
        section.remove();
        render();
      } catch (error) {
        title.textContent = String(error);
        close.disabled = false;
      }
    }
    tools.append(close);
    const frame = document.createElement('iframe');
    frame.title = label;
    frame.src = '/agents/' + id;
    frames.set(id, frame);
    panes.set(id, section);
    section.append(tools, frame);
    return section;
  }
  const dividers = new Map<Layout, HTMLElement>();
  const bounds = new Map<Layout, { x: number; y: number; width: number; height: number }>();
  function position(node: Layout, x = 0, y = 0, width = 100, height = 100) {
    bounds.set(node, { x, y, width, height });
    if ('id' in node) {
      const element = panes.get(node.id)!;
      Object.assign(element.style, {
        left: x + '%',
        top: y + '%',
        width: width + '%',
        height: height + '%',
      });
      return;
    }
    const divider = dividers.get(node)!;
    const horizontal = node.direction === 'horizontal';
    Object.assign(
      divider.style,
      horizontal
        ? {
            left: x + width * node.ratio + '%',
            top: y + '%',
            width: '6px',
            height: height + '%',
            transform: 'translateX(-3px)',
          }
        : {
            left: x + '%',
            top: y + height * node.ratio + '%',
            width: width + '%',
            height: '6px',
            transform: 'translateY(-3px)',
          },
    );
    divider.setAttribute('aria-valuenow', String(Math.round(node.ratio * 100)));
    if (horizontal) {
      position(node.first, x, y, width * node.ratio, height);
      position(node.second, x + width * node.ratio, y, width * (1 - node.ratio), height);
    } else {
      position(node.first, x, y, width, height * node.ratio);
      position(node.second, x, y + height * node.ratio, width, height * (1 - node.ratio));
    }
  }
  function tree(node: Layout) {
    if ('id' in node) {
      const element = pane(node.id);
      // Keep existing frames connected so splitting preserves drafts and terminal state.
      if (!element.isConnected) host.append(element);
      return;
    }
    const divider = document.createElement('div');
    divider.className = 'agent-divider ' + node.direction;
    divider.tabIndex = 0;
    divider.setAttribute('role', 'separator');
    divider.setAttribute('aria-label', 'Agent split');
    divider.setAttribute(
      'aria-orientation',
      node.direction === 'horizontal' ? 'vertical' : 'horizontal',
    );
    divider.setAttribute('aria-valuemin', '15');
    divider.setAttribute('aria-valuemax', '85');
    const resize = (ratio: number) => {
      node.ratio = Math.max(0.15, Math.min(0.85, ratio));
      position(layout);
    };
    divider.onpointerdown = (event) => {
      if (event.button !== 0) return;
      event.preventDefault();
      divider.setPointerCapture(event.pointerId);
      document.body.classList.add('splitting-agents');
    };
    divider.onpointermove = (event) => {
      if (!divider.hasPointerCapture(event.pointerId)) return;
      const rect = host.getBoundingClientRect(),
        box = bounds.get(node)!;
      resize(
        node.direction === 'horizontal'
          ? (((event.clientX - rect.left) / rect.width) * 100 - box.x) / box.width
          : (((event.clientY - rect.top) / rect.height) * 100 - box.y) / box.height,
      );
    };
    divider.onlostpointercapture = () => {
      document.body.classList.remove('splitting-agents');
      save();
    };
    divider.onkeydown = (event) => {
      const keys =
        node.direction === 'horizontal' ? ['ArrowLeft', 'ArrowRight'] : ['ArrowUp', 'ArrowDown'];
      if (!keys.includes(event.key)) return;
      event.preventDefault();
      event.stopPropagation();
      resize(node.ratio + (event.key === keys[0] ? -0.05 : 0.05));
      save();
    };
    dividers.set(node, divider);
    host.append(divider);
    tree(node.first);
    tree(node.second);
  }
  function render() {
    for (const divider of dividers.values()) divider.remove();
    dividers.clear();
    bounds.clear();
    tree(layout);
    position(layout);
    for (const pane of panes.values())
      for (const button of pane.querySelectorAll<HTMLButtonElement>('[data-split]'))
        button.disabled = panes.size >= MAX_PANES;
    save();
  }
  render();
}
