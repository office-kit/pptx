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
  let mode: 'region' | 'text' = 'region';
  let selection: Selection | undefined;
  let sending = false;
  let pendingReview: { prompt: string; slide: number | null } | undefined;
  let start: { x: number; y: number; target: Element | null } | undefined;
  let textTarget: Element | null = null;
  const tools = document.createElement('div');
  tools.className = 'slide-edit-tools';
  tools.innerHTML =
    '<button type="button" data-undo aria-label="Undo" disabled title="Undo (⌘/Ctrl+Z)"><svg width="18" height="18" viewBox="0 0 16 16" aria-hidden="true"><path fill="currentColor" d="M12.8332 9.66691C12.8332 9.25106 12.7516 8.8392 12.5925 8.455C12.4333 8.07081 12.2 7.7217 11.9059 7.42766C11.6119 7.13361 11.2628 6.90027 10.8786 6.74113C10.4944 6.58199 10.0825 6.50041 9.66667 6.50041H3.8737L6.35368 8.9799C6.54889 9.17511 6.54878 9.49166 6.35368 9.68693C6.15842 9.88219 5.84191 9.88219 5.64665 9.68693L2.31315 6.35393C2.11792 6.15869 2.11797 5.84216 2.31315 5.64689L5.64665 2.3134C5.84192 2.11822 6.15845 2.11816 6.35368 2.3134C6.54883 2.50864 6.54883 2.82519 6.35368 3.02043L3.8737 5.50041H9.66667C10.2138 5.50041 10.7559 5.60791 11.2614 5.8173C11.7669 6.0267 12.2261 6.33375 12.613 6.72062C12.9998 7.1075 13.3069 7.56672 13.5163 8.07219C13.7257 8.57771 13.8332 9.11974 13.8332 9.66691C13.8332 10.2141 13.7257 10.7561 13.5163 11.2616C13.3069 11.7671 12.9998 12.2263 12.613 12.6132C12.2261 13.0001 11.7669 13.3071 11.2614 13.5165C10.7559 13.7259 10.2138 13.8334 9.66667 13.8334H7.33317C7.05711 13.8333 6.83317 13.6095 6.83317 13.3334C6.83326 13.0574 7.05716 12.8335 7.33317 12.8334H9.66667C10.0825 12.8334 10.4944 12.7518 10.8786 12.5927C11.2628 12.4336 11.6119 12.2002 11.9059 11.9062C12.2 11.6121 12.4333 11.263 12.5925 10.8788C12.7516 10.4946 12.8332 10.0828 12.8332 9.66691Z"/></svg></button><button type="button" data-redo aria-label="Redo" disabled title="Redo (⌘/Ctrl+Shift+Z)"><svg width="18" height="18" viewBox="0 0 16 16" aria-hidden="true"><path fill="currentColor" d="M3.16683 9.66691C3.16683 9.25106 3.24841 8.8392 3.40755 8.455C3.56669 8.07081 3.80002 7.7217 4.09407 7.42766C4.38812 7.13361 4.73722 6.90027 5.12142 6.74113C5.50561 6.58199 5.91748 6.50041 6.33333 6.50041H12.1263L9.64632 8.9799C9.45111 9.17511 9.45122 9.49166 9.64632 9.68693C9.84158 9.88219 10.1581 9.88219 10.3533 9.68693L13.6868 6.35393C13.8821 6.15869 13.882 5.84216 13.6868 5.64689L10.3533 2.3134C10.1581 2.11822 9.84155 2.11816 9.64632 2.3134C9.45117 2.50864 9.45117 2.82519 9.64632 3.02043L12.1263 5.50041H6.33333C5.78616 5.50041 5.24413 5.60791 4.7386 5.8173C4.23314 6.0267 3.77392 6.33375 3.38704 6.72062C3.00017 7.1075 2.69311 7.56672 2.48372 8.07219C2.27433 8.57771 2.16683 9.11974 2.16683 9.66691C2.16683 10.2141 2.27433 10.7561 2.48372 11.2616C2.69311 11.7671 3.00017 12.2263 3.38704 12.6132C3.77392 13.0001 4.23314 13.3071 4.7386 13.5165C5.24413 13.7259 5.78616 13.8334 6.33333 13.8334H8.66683C8.94289 13.8333 9.16683 13.6095 9.16683 13.3334C9.16674 13.0574 8.94284 12.8335 8.66683 12.8334H6.33333C5.91748 12.8334 5.50561 12.7518 5.12142 12.5927C4.73722 12.4336 4.38812 12.2002 4.09407 11.9062C3.80002 11.6121 3.56669 11.263 3.40755 10.8788C3.24841 10.4946 3.16683 10.0828 3.16683 9.66691Z"/></svg></button><span>Click to select · Double-click text to edit · Drag to select an area</span><small role="status" data-history-status></small>';
  document.querySelector('main')!.prepend(tools);
  const outline = document.createElement('div');
  outline.className = 'slide-selection';
  outline.hidden = true;
  const panel = document.createElement('form');
  panel.className = 'slide-edit-panel';
  panel.hidden = true;
  panel.setAttribute('aria-label', 'Edit selection');
  panel.innerHTML =
    '<div class="selection-heading"><strong></strong><button type="button" data-cancel aria-label="Cancel selection">×</button></div><textarea aria-label="Selection edit" maxlength="12000" required></textarea><div class="selection-actions"><select aria-label="Editing agent"></select><button type="button" data-review hidden>Retry visual review</button><button type="submit">Apply with AI</button></div><small role="status"></small>';
  const hover = document.createElement('div');
  hover.className = 'slide-hover';
  hover.hidden = true;
  const textInput = document.createElement('textarea');
  textInput.className = 'slide-text-input';
  textInput.setAttribute('aria-label', 'Edit slide text');
  textInput.maxLength = 12000;
  textInput.hidden = true;
  document.body.append(outline, hover, panel, textInput);
  const input = panel.querySelector('textarea')!;
  const agents = panel.querySelector('select')!;
  const status = panel.querySelector('small')!;
  const retry = panel.querySelector<HTMLButtonElement>('[data-review]')!;
  const undo = tools.querySelector<HTMLButtonElement>('[data-undo]')!;
  const redo = tools.querySelector<HTMLButtonElement>('[data-redo]')!;
  const historyStatus = tools.querySelector<HTMLElement>('[data-history-status]')!;
  tools.append(retry);
  function setStatus(message: string) {
    (mode === 'text' || panel.hidden ? historyStatus : status).textContent = message;
  }
  function close() {
    selection = undefined;
    start = undefined;
    outline.hidden = panel.hidden = textInput.hidden = hover.hidden = true;
    textTarget = null;
    mode = 'region';
  }
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
    Object.assign(textInput.style, {
      left: left + 'px',
      top: top + 'px',
      width: Math.max(80, selection.width * rect.width) + 'px',
      height: Math.max(40, selection.height * rect.height) + 'px',
    });
    Object.assign(panel.style, {
      left: Math.max(8, Math.min(innerWidth - panel.offsetWidth - 8, left)) + 'px',
      top:
        Math.max(
          8,
          Math.min(
            innerHeight - panel.offsetHeight - 8,
            top + Math.max(mode === 'text' ? 40 : 0, selection.height * rect.height) + 8,
          ),
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
      'Selected area · Slide ' + (selection.focus.slide! + 1);
    input.value = '';
    input.hidden = mode === 'text';
    input.required = mode !== 'text';
    textInput.hidden = mode !== 'text';
    textInput.value = text;
    input.placeholder = 'Move this down a little…';
    panel.hidden = mode === 'text';
    setStatus(
      mode === 'text'
        ? 'Shift+Enter to save · Enter for newline · Esc to cancel.'
        : 'Shift+Enter to send · Enter for newline.',
    );
    outline.hidden = false;
    position();
    if (mode === 'text') {
      if (textTarget) {
        const style = getComputedStyle(textTarget);
        const scale =
          textTarget.getBoundingClientRect().width /
          (textTarget instanceof HTMLElement ? textTarget.offsetWidth || 1 : 1);
        textInput.style.font = style.font;
        textInput.style.fontSize = Math.max(14, parseFloat(style.fontSize) * scale) + 'px';
        textInput.style.textAlign = style.textAlign;
      }
      textInput.focus();
      textInput.select();
    }
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
  function objectAt(target: EventTarget | null) {
    if (!(target instanceof Element)) return null;
    const object =
      target.closest('[data-pptx-shape-name]') ??
      target.closest('p, image, rect, path, ellipse, polygon, line');
    if (!object) return null;
    const box = object.getBoundingClientRect(),
      rect = slide.getBoundingClientRect();
    if (box.width >= rect.width * 0.98 && box.height >= rect.height * 0.98) return null;
    return object;
  }
  function selectObject(target: Element, editing = false) {
    close();
    mode = editing ? 'text' : 'region';
    textTarget = editing
      ? target
      : (target.querySelector('p') ?? (target.matches('p') ? target : null));
    const rect = slide.getBoundingClientRect(),
      box = target.getBoundingClientRect();
    selection = {
      x: (box.left - rect.left) / rect.width,
      y: (box.top - rect.top) / rect.height,
      width: box.width / rect.width,
      height: box.height / rect.height,
      text: '',
      focus: focus(),
    };
    show(editing ? (target.textContent ?? '') : selectedText());
  }
  const interactive = () => !sending && !document.body.classList.contains('presenting');
  root.addEventListener('pointerdown', (event) => {
    if (!(event instanceof PointerEvent) || event.button !== 0 || !interactive()) return;
    event.preventDefault();
    slide.focus({ preventScroll: true });
    close();
    const rect = slide.getBoundingClientRect();
    start = {
      x: (event.clientX - rect.left) / rect.width,
      y: (event.clientY - rect.top) / rect.height,
      target: objectAt(event.target),
    };
  });
  root.addEventListener('pointermove', (event) => {
    if (!(event instanceof PointerEvent) || start || !interactive() || mode === 'text') return;
    const target = objectAt(event.target);
    hover.hidden = !target;
    if (target) {
      const box = target.getBoundingClientRect();
      Object.assign(hover.style, {
        left: box.left + 'px',
        top: box.top + 'px',
        width: box.width + 'px',
        height: box.height + 'px',
      });
    }
  });
  slide.addEventListener('pointerleave', () => {
    hover.hidden = true;
  });
  window.addEventListener('pointermove', (event) => {
    if (!start) return;
    const rect = slide.getBoundingClientRect();
    const x = Math.max(0, Math.min(1, (event.clientX - rect.left) / rect.width));
    const y = Math.max(0, Math.min(1, (event.clientY - rect.top) / rect.height));
    if (Math.hypot((x - start.x) * rect.width, (y - start.y) * rect.height) < 5) return;
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
  });
  window.addEventListener('pointerup', () => {
    if (!start) return;
    const target = start.target;
    start = undefined;
    if (selection && selection.width > 0.005 && selection.height > 0.005) show(selectedText());
    else if (target) selectObject(target);
    else close();
  });
  window.addEventListener('pointercancel', close);
  root.addEventListener('dblclick', (event) => {
    if (!interactive()) return;
    const target = event.target instanceof Element ? event.target.closest('p') : null;
    if (!target?.textContent?.trim()) return;
    event.preventDefault();
    selectObject(target, true);
  });
  slide.tabIndex = 0;
  slide.setAttribute(
    'aria-label',
    'Slide canvas. Click to select, double-click text to edit, or drag to select an area.',
  );
  slide.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' && textTarget && interactive()) {
      event.preventDefault();
      selectObject(textTarget, true);
    }
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
      setStatus(error instanceof Error ? error.message : String(error));
    } finally {
      sending = false;
      textInput.readOnly = false;
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
        'Instruction: ' +
        input.value +
        '\nLocate the corresponding TSX before editing. Change only the selected instance, preserving other slides and shared components unless explicitly requested. Verify the updated preview.';
      status.textContent = 'Sending…';
      await send(message, agents.value, selection.focus);
      status.textContent = 'Sent · follow progress in the agent pane.';
    });
  };
  input.onkeydown = (event) => {
    if (
      event.key === 'Enter' &&
      (event.shiftKey || event.metaKey || event.ctrlKey) &&
      !event.isComposing
    ) {
      event.preventDefault();
      panel.requestSubmit();
    }
  };
  const saveText = () =>
    void run(async () => {
      if (!selection) return;
      const edited = selection;
      textInput.readOnly = true;
      setStatus('Saving…');
      const response = await fetch('/text-edit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          ...selection.focus,
          before: selection.text,
          after: textInput.value,
        }),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error);
      close();
      setStatus('Text saved. Sending visual review…');
      if (result.reviewError) throw new Error(result.reviewError);
      if (result.review) {
        pendingReview = { prompt: result.review.prompt, slide: edited.focus.slide };
        retry.hidden = false;
        await sendReview();
      }
      setStatus('Text saved · the agent is reviewing the updated slide.');
    });
  textInput.onkeydown = (event) => {
    if (
      event.key === 'Enter' &&
      (event.shiftKey || event.metaKey || event.ctrlKey) &&
      !event.isComposing
    ) {
      event.preventDefault();
      saveText();
    }
  };
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
      setStatus('Visual review sent · follow progress in the agent pane.');
    });
  let historyBusy = false;
  async function refreshHistory() {
    try {
      const state = await (await fetch('/history')).json();
      undo.disabled = historyBusy || state.busy || !state.undo;
      redo.disabled = historyBusy || state.busy || !state.redo;
      undo.title = state.undo ? 'Undo: ' + state.undo + ' (⌘/Ctrl+Z)' : 'Nothing to undo';
      redo.title = state.redo ? 'Redo: ' + state.redo + ' (⌘/Ctrl+Shift+Z)' : 'Nothing to redo';
      if (state.error) historyStatus.textContent = state.error;
    } catch {
      historyStatus.textContent = 'History unavailable. Check the dev server connection.';
    }
  }
  async function moveHistory(direction: 'undo' | 'redo') {
    if (historyBusy || sending) return;
    historyBusy = true;
    undo.disabled = redo.disabled = true;
    close();
    try {
      const response = await fetch('/history/' + direction, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: '{}',
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error);
      historyStatus.textContent = result.buildError
        ? 'Source restored; preview has a build error.'
        : direction === 'undo'
          ? 'Undone'
          : 'Redone';
    } catch (cause) {
      historyStatus.textContent = cause instanceof Error ? cause.message : String(cause);
    } finally {
      historyBusy = false;
      await refreshHistory();
    }
  }
  undo.onclick = () => void moveHistory('undo');
  redo.onclick = () => void moveHistory('redo');
  window.addEventListener('agent-history', () => void refreshHistory());
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !sending) close();
    const target = event.composedPath()[0];
    if (
      event.isComposing ||
      (target instanceof Element &&
        target.closest('input, textarea, select, [contenteditable], iframe'))
    )
      return;
    if (
      (event.metaKey || event.ctrlKey) &&
      !event.altKey &&
      (event.key.toLowerCase() === 'z' || event.key.toLowerCase() === 'y')
    ) {
      event.preventDefault();
      const direction = event.shiftKey || event.key.toLowerCase() === 'y' ? 'redo' : 'undo';
      if (!(direction === 'undo' ? undo : redo).disabled) void moveHistory(direction);
    }
  });
  window.addEventListener('agent-focus', () =>
    queueMicrotask(() => {
      if (selection && JSON.stringify(selection.focus) !== JSON.stringify(focus())) {
        if (!sending) close();
        else if (mode === 'text') textInput.hidden = outline.hidden = true;
      }
    }),
  );
  new MutationObserver(() => {
    if (document.body.classList.contains('presenting')) close();
  }).observe(document.body, { attributes: true, attributeFilter: ['class'] });
  new ResizeObserver(position).observe(slide);
  stage.addEventListener('scroll', position);
  window.addEventListener('resize', position);
  void refreshHistory();
}
