interface Focus {
  slide: number | null;
  revision: number;
}
interface Selection {
  x: number;
  y: number;
  width: number;
  height: number;
  text: string;
  focus: Focus;
}

export function mountEditor(frames: Map<string, HTMLIFrameElement>) {
  const slide = document.getElementById('slide')!;
  const stage = document.getElementById('stage')!;
  const root = slide.shadowRoot!;
  const focus = (): Focus => JSON.parse(document.getElementById('chat-context')!.dataset.focus!);
  let mode: 'region' | 'text' | null = null;
  let selection: Selection | undefined;
  let sending = false;
  let pendingReview: { prompt: string; slide: number | null } | undefined;
  let lastRequest: { message: string; agent: string } | undefined;
  let start: { x: number; y: number } | undefined;
  const tools = document.createElement('div');
  tools.className = 'slide-edit-tools';
  tools.innerHTML =
    '<button type="button" data-mode="region" aria-pressed="false">✦ Select area</button><button type="button" data-mode="text" aria-pressed="false">Edit text</button><button type="button" data-undo hidden>Undo text</button><button type="button" data-undo-ai hidden>Undo with AI</button>';
  document.querySelector('main')!.prepend(tools);
  const outline = document.createElement('div');
  outline.className = 'slide-selection';
  outline.hidden = true;
  const panel = document.createElement('form');
  panel.className = 'slide-edit-panel';
  panel.hidden = true;
  panel.setAttribute('aria-label', 'Edit selection');
  panel.innerHTML =
    '<div class="selection-heading"><strong></strong><button type="button" data-cancel aria-label="Cancel selection">×</button></div><textarea aria-label="Selection edit" maxlength="12000" required></textarea><div class="selection-actions"><select aria-label="Editing agent"></select><button type="button" data-review hidden>Retry visual review</button><button type="button" data-save>Save text</button><button type="submit">Apply with AI</button></div><small role="status"></small>';
  document.body.append(outline, panel);
  const input = panel.querySelector('textarea')!;
  const agents = panel.querySelector('select')!;
  const status = panel.querySelector('small')!;
  const save = panel.querySelector<HTMLButtonElement>('[data-save]')!;
  const retry = panel.querySelector<HTMLButtonElement>('[data-review]')!;
  const undo = tools.querySelector<HTMLButtonElement>('[data-undo]')!;
  const undoAI = tools.querySelector<HTMLButtonElement>('[data-undo-ai]')!;
  function close() {
    selection = undefined;
    start = undefined;
    outline.hidden = panel.hidden = true;
  }
  function setMode(next: typeof mode) {
    mode = next;
    close();
    for (const button of tools.querySelectorAll<HTMLButtonElement>('[data-mode]'))
      button.setAttribute('aria-pressed', String(button.dataset.mode === mode));
    attachOverlay();
  }
  for (const button of tools.querySelectorAll<HTMLButtonElement>('[data-mode]'))
    button.onclick = () =>
      setMode(mode === button.dataset.mode ? null : (button.dataset.mode as 'text' | 'region'));
  panel.querySelector<HTMLButtonElement>('[data-cancel]')!.onclick = close;
  function position() {
    if (!selection) return;
    const rect = slide.getBoundingClientRect();
    const left = rect.left + selection.x * rect.width,
      top = rect.top + selection.y * rect.height;
    Object.assign(outline.style, {
      left: left + 'px',
      top: top + 'px',
      width: selection.width * rect.width + 'px',
      height: selection.height * rect.height + 'px',
    });
    Object.assign(panel.style, {
      left: Math.max(8, Math.min(innerWidth - panel.offsetWidth - 8, left)) + 'px',
      top:
        Math.max(
          8,
          Math.min(innerHeight - panel.offsetHeight - 8, top + selection.height * rect.height + 8),
        ) + 'px',
    });
  }
  function show(text: string) {
    if (!selection) return;
    selection.text = text;
    const previous = agents.value;
    agents.replaceChildren(...[...frames].map(([id, frame]) => new Option(frame.title, id)));
    if (frames.has(previous)) agents.value = previous;
    panel.querySelector('strong')!.textContent =
      mode === 'text'
        ? 'Edit text · Slide ' + (selection.focus.slide! + 1)
        : 'Selected area · Slide ' + (selection.focus.slide! + 1);
    input.value = mode === 'text' ? text : '';
    input.placeholder = mode === 'text' ? 'Replacement text…' : 'Move this down a little…';
    save.hidden = mode !== 'text';
    status.textContent =
      mode === 'text'
        ? 'Save an exact source literal, or apply with AI for computed text.'
        : 'Only this area is edited unless you ask for a broader change.';
    panel.hidden = outline.hidden = false;
    position();
    input.focus();
  }
  function selectedText() {
    if (!selection) return '';
    const rect = slide.getBoundingClientRect();
    return [...root.querySelectorAll('p')]
      .filter((p) => {
        const box = p.getBoundingClientRect();
        return (
          box.right > rect.left + selection!.x * rect.width &&
          box.left < rect.left + (selection!.x + selection!.width) * rect.width &&
          box.bottom > rect.top + selection!.y * rect.height &&
          box.top < rect.top + (selection!.y + selection!.height) * rect.height
        );
      })
      .map((p) => p.textContent)
      .join('\n')
      .slice(0, 8000);
  }
  function attachOverlay() {
    root.querySelector('[data-selection-overlay]')?.remove();
    slide.style.position = 'relative';
    const overlay = document.createElement('div');
    overlay.dataset.selectionOverlay = '';
    Object.assign(overlay.style, {
      position: 'absolute',
      inset: '0',
      cursor: 'crosshair',
      touchAction: 'none',
      userSelect: 'none',
      pointerEvents: mode === 'region' ? 'auto' : 'none',
    });
    overlay.onpointerdown = (event) => {
      if (event.button !== 0 || sending) return;
      event.preventDefault();
      close();
      const rect = slide.getBoundingClientRect();
      start = {
        x: (event.clientX - rect.left) / rect.width,
        y: (event.clientY - rect.top) / rect.height,
      };
      overlay.setPointerCapture(event.pointerId);
    };
    overlay.onpointermove = (event) => {
      if (!start) return;
      const rect = slide.getBoundingClientRect();
      const x = Math.max(0, Math.min(1, (event.clientX - rect.left) / rect.width));
      const y = Math.max(0, Math.min(1, (event.clientY - rect.top) / rect.height));
      selection = {
        x: Math.min(start.x, x),
        y: Math.min(start.y, y),
        width: Math.abs(x - start.x),
        height: Math.abs(y - start.y),
        text: '',
        focus: focus(),
      };
      outline.hidden = false;
      position();
    };
    overlay.onpointerup = () => {
      start = undefined;
      if (selection && selection.width > 0.005 && selection.height > 0.005) show(selectedText());
      else close();
    };
    overlay.onpointercancel = close;
    root.append(overlay);
  }
  root.addEventListener('click', (event) => {
    if (mode !== 'text' || sending) return;
    const target = event.target instanceof Element ? event.target.closest('p') : null;
    if (!target?.textContent?.trim()) return;
    event.preventDefault();
    const rect = slide.getBoundingClientRect(),
      box = target.getBoundingClientRect();
    selection = {
      x: (box.left - rect.left) / rect.width,
      y: (box.top - rect.top) / rect.height,
      width: box.width / rect.width,
      height: box.height / rect.height,
      text: target.textContent,
      focus: focus(),
    };
    show(target.textContent);
  });
  function send(message: string, agent: string, context: Focus) {
    const frame = frames.get(agent);
    if (!frame?.contentWindow) return Promise.reject(new Error('Select an available agent.'));
    const target = frame.contentWindow;
    return new Promise<void>((resolve, reject) => {
      const id = crypto.randomUUID();
      const timer = setTimeout(() => {
        window.removeEventListener('message', listener);
        reject(new Error('Agent did not acknowledge the request. Check its pane before retrying.'));
      }, 30000);
      const listener = (event: MessageEvent) => {
        if (
          event.origin !== location.origin ||
          event.source !== frame.contentWindow ||
          event.data?.type !== 'inline-result' ||
          event.data.id !== id
        )
          return;
        clearTimeout(timer);
        window.removeEventListener('message', listener);
        if (event.data.error) reject(new Error(event.data.error));
        else resolve();
      };
      window.addEventListener('message', listener);
      target.postMessage({ type: 'inline-edit', id, message, focus: context }, location.origin);
    });
  }
  async function run(action: () => Promise<void>) {
    if (sending) return;
    sending = true;
    for (const button of panel.querySelectorAll('button')) button.disabled = true;
    try {
      await action();
    } catch (error) {
      status.textContent = error instanceof Error ? error.message : String(error);
    } finally {
      sending = false;
      for (const button of panel.querySelectorAll('button')) button.disabled = false;
      position();
    }
  }
  panel.onsubmit = (event) => {
    event.preventDefault();
    void run(async () => {
      if (!selection) return;
      if (JSON.stringify(selection.focus) !== JSON.stringify(focus()))
        throw new Error('Preview changed. Select the area again.');
      const message =
        'Edit the selected area of slide ' +
        (selection.focus.slide! + 1) +
        '. Bounds are fractions of the full slide, origin top-left: ' +
        JSON.stringify({
          x: selection.x,
          y: selection.y,
          width: selection.width,
          height: selection.height,
        }) +
        '\nText intersecting the selection: ' +
        JSON.stringify(selection.text) +
        '\n' +
        (mode === 'text'
          ? 'Replace this text with exactly: ' + JSON.stringify(input.value)
          : 'Instruction: ' + input.value) +
        '\nLocate the corresponding TSX before editing. Change only the selected instance, preserving other slides and shared components unless explicitly requested. Verify the updated preview.';
      status.textContent = 'Sending…';
      await send(message, agents.value, selection.focus);
      lastRequest = { message, agent: agents.value };
      undoAI.hidden = false;
      status.textContent = 'Sent · follow progress in the agent pane.';
    });
  };
  input.onkeydown = (event) => {
    if (event.key === 'Enter' && (event.metaKey || event.ctrlKey) && !event.isComposing) {
      event.preventDefault();
      panel.requestSubmit();
    }
  };
  save.onclick = () =>
    void run(async () => {
      if (!selection) return;
      status.textContent = 'Saving and capturing the updated slide…';
      const response = await fetch('/text-edit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ...selection.focus, before: selection.text, after: input.value }),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error);
      undo.hidden = false;
      status.textContent = 'Text saved. Sending visual review…';
      if (result.reviewError) throw new Error(result.reviewError);
      if (result.review) {
        pendingReview = { prompt: result.review.prompt, slide: selection.focus.slide };
        retry.hidden = false;
        await sendReview();
      }
      status.textContent = 'Text saved · the agent is reviewing the updated slide.';
    });
  async function sendReview() {
    if (!pendingReview) return;
    const state = await (await fetch('/state')).json();
    await send(pendingReview.prompt, agents.value, {
      slide: pendingReview.slide,
      revision: state.revision,
    });
    pendingReview = undefined;
    retry.hidden = true;
  }
  retry.onclick = () =>
    void run(async () => {
      await sendReview();
      status.textContent = 'Visual review sent · follow progress in the agent pane.';
    });
  undo.onclick = () =>
    void run(async () => {
      const response = await fetch('/text-edit/undo', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: '{}',
      });
      const result = await response.json();
      if (!response.ok) {
        panel.hidden = false;
        throw new Error(result.error);
      }
      undo.hidden = true;
      panel.hidden = false;
      if (result.reviewError) throw new Error(result.reviewError);
      if (result.review) {
        pendingReview = { prompt: result.review.prompt, slide: focus().slide };
        retry.hidden = false;
        await sendReview();
      }
      close();
    });
  undoAI.onclick = () =>
    void run(async () => {
      if (!lastRequest) return;
      panel.hidden = false;
      await send(
        'Undo only the changes from this previous request, preserving any newer or unrelated work. Inspect source before reversing anything. Previous request: ' +
          lastRequest.message,
        lastRequest.agent,
        focus(),
      );
      status.textContent = 'Undo requested · follow progress in the agent pane.';
      undoAI.hidden = true;
    });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !sending) setMode(null);
  });
  window.addEventListener('agent-focus', () =>
    queueMicrotask(() => {
      if (selection && JSON.stringify(selection.focus) !== JSON.stringify(focus()) && !sending)
        close();
      attachOverlay();
    }),
  );
  new MutationObserver(() => {
    if (document.body.classList.contains('presenting')) setMode(null);
  }).observe(document.body, { attributes: true, attributeFilter: ['class'] });
  new ResizeObserver(position).observe(slide);
  stage.addEventListener('scroll', position);
  window.addEventListener('resize', position);
  attachOverlay();
}
