import type { EditorLink, EditorLinkAction } from './editor.ts';

export function openLinkDialog(
  initial: EditorLink | null,
  slides: Array<{ key: string; title: string }>,
  apply: (link: EditorLink | null, displayText?: string) => Promise<boolean>,
  restoreFocus: () => void,
  selectedText?: string,
  customShows: Array<{ id: number; name: string }> = [],
) {
  const dialog = document.createElement('dialog');
  dialog.className = 'section-dialog link-dialog';
  dialog.setAttribute('aria-label', initial ? 'Edit Hyperlink' : 'Insert Hyperlink');
  dialog.innerHTML = `<form><h2>${initial ? 'Edit' : 'Insert'} Hyperlink</h2>
    <label>Text to Display: <input value="&lt;&lt;Selection in Document&gt;&gt;" disabled></label>
    <button type="button" data-tip>ScreenTip...</button>
    <div role="tablist" aria-label="Link destination">
      <button type="button" role="tab" data-tab="web">Web Page or File</button>
      <button type="button" role="tab" data-tab="document">This Document</button>
      <button type="button" role="tab" data-tab="email">Email Address</button>
    </div>
    <section data-panel="web" role="tabpanel"><p>Link to an existing file or web page.</p><label>Address: <input name="address"></label></section>
    <section data-panel="document" role="tabpanel"><label>Select a place in this document: <select name="destination" size="7"></select></label><label data-return-label hidden><input type="checkbox" data-return>Show and return</label></section>
    <section data-panel="email" role="tabpanel"><label>Email address: <input name="email"></label><label>Subject: <input name="subject"></label></section>
    <p role="alert"></p><footer><button type="button" data-remove>Remove Link</button><button type="button" data-cancel>Cancel</button><button type="submit">OK</button></footer></form>`;
  const display = dialog.querySelector<HTMLInputElement>('input')!;
  if (selectedText !== undefined) {
    display.value = selectedText;
    display.disabled = false;
  }
  const field = (name: string) => dialog.querySelector<HTMLInputElement>(`[name="${name}"]`)!;
  const destination = dialog.querySelector<HTMLSelectElement>('[name="destination"]')!;
  for (const [value, label] of [
    ['firstSlide', 'First Slide'],
    ['lastSlide', 'Last Slide'],
    ['nextSlide', 'Next Slide'],
    ['prevSlide', 'Previous Slide'],
    ...(initial?.action.kind === 'lastSlideViewed'
      ? [['lastSlideViewed', 'Last Slide Viewed']]
      : []),
    ...(initial?.action.kind === 'endShow' ? [['endShow', 'End Show']] : []),
  ])
    destination.add(new Option(label, value));
  slides.forEach((slide, index) =>
    destination.add(
      new Option(`${index + 1}. ${slide.title || 'Slide ' + (index + 1)}`, slide.key),
    ),
  );
  customShows.forEach((show) => destination.add(new Option(show.name, `customShow:${show.id}`)));
  const returnLabel = dialog.querySelector<HTMLElement>('[data-return-label]')!;
  const returnBox = dialog.querySelector<HTMLInputElement>('[data-return]')!;
  returnBox.checked = initial?.action.kind === 'customShow' && initial.action.showAndReturn;
  const updateReturn = () => {
    returnLabel.hidden = !destination.value.startsWith('customShow:');
  };
  destination.onchange = updateReturn;
  let tab = 'web';
  let tooltip = initial?.tooltip ?? null;
  if (initial?.action.kind === 'url') {
    field('address').value = initial.action.url;
    // Preserve arbitrary existing mailto parameters unless the email fields are edited.
    if (/^mailto:/i.test(initial.action.url)) {
      tab = 'email';
      const [address, query] = initial.action.url.slice(7).split('?');
      try {
        field('email').value = decodeURIComponent(address!);
      } catch {
        field('email').value = address!;
      }
      field('subject').value = new URLSearchParams(query).get('subject') ?? '';
    }
  } else if (initial) {
    tab = 'document';
    destination.value =
      initial.action.kind === 'slide'
        ? initial.action.slide
        : initial.action.kind === 'customShow'
          ? `customShow:${initial.action.id}`
          : initial.action.kind;
  }
  updateReturn();
  let emailEdited = false;
  for (const name of ['email', 'subject'])
    field(name).addEventListener('input', () => {
      emailEdited = true;
    });
  const tabs = [...dialog.querySelectorAll<HTMLButtonElement>('[data-tab]')];
  const update = () => {
    tabs.forEach((button) => {
      const active = button.dataset.tab === tab;
      button.setAttribute('aria-selected', String(active));
      button.tabIndex = active ? 0 : -1;
      button.id = `link-tab-${button.dataset.tab}`;
      button.setAttribute('aria-controls', `link-panel-${button.dataset.tab}`);
    });
    dialog.querySelectorAll<HTMLElement>('[data-panel]').forEach((panel) => {
      panel.hidden = panel.dataset.panel !== tab;
      panel.id = `link-panel-${panel.dataset.panel}`;
      panel.setAttribute('aria-labelledby', `link-tab-${panel.dataset.panel}`);
    });
  };
  tabs.forEach((button, index) => {
    button.onclick = () => {
      tab = button.dataset.tab!;
      update();
    };
    button.onkeydown = (event) => {
      if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
      event.preventDefault();
      const next =
        event.key === 'Home'
          ? 0
          : event.key === 'End'
            ? 2
            : (index + (event.key === 'ArrowRight' ? 1 : 2)) % 3;
      tabs[next]!.click();
      tabs[next]!.focus();
    };
  });
  const close = () => {
    dialog.close();
    dialog.remove();
    restoreFocus();
  };
  let submitting = false;
  dialog.addEventListener('cancel', (event) => {
    event.preventDefault();
    if (!submitting) close();
  });
  dialog.querySelector<HTMLButtonElement>('[data-cancel]')!.onclick = close;
  const save = async (link: EditorLink | null) => {
    if (submitting) return;
    submitting = true;
    dialog.querySelectorAll<HTMLButtonElement>('button').forEach((button) => {
      button.disabled = true;
    });
    dialog.close();
    try {
      if (await apply(link, link && selectedText !== undefined ? display.value : undefined)) {
        close();
        return;
      }
    } catch (error) {
      dialog.querySelector('[role="alert"]')!.textContent =
        error instanceof Error ? error.message : 'Unable to save link.';
    }
    submitting = false;
    dialog.querySelectorAll<HTMLButtonElement>('button').forEach((button) => {
      button.disabled = false;
    });
    dialog.showModal();
  };
  const remove = dialog.querySelector<HTMLButtonElement>('[data-remove]')!;
  remove.hidden = !initial;
  remove.onclick = () => void save(null);
  dialog.querySelector<HTMLButtonElement>('[data-tip]')!.onclick = () => {
    const tip = document.createElement('dialog');
    tip.className = 'section-dialog';
    tip.setAttribute('aria-label', 'Set Hyperlink ScreenTip');
    tip.innerHTML =
      '<form><h2>Set Hyperlink ScreenTip</h2><label>ScreenTip text: <input></label><footer><button type="button">Cancel</button><button type="submit">OK</button></footer></form>';
    const input = tip.querySelector('input')!;
    input.value = tooltip ?? '';
    const closeTip = () => {
      tip.close();
      tip.remove();
      dialog.querySelector<HTMLButtonElement>('[data-tip]')!.focus();
    };
    tip.addEventListener('cancel', (event) => {
      event.preventDefault();
      closeTip();
    });
    tip.querySelector('button')!.onclick = closeTip;
    tip.querySelector('form')!.onsubmit = (event) => {
      event.preventDefault();
      tooltip = input.value || null;
      closeTip();
    };
    document.body.append(tip);
    tip.showModal();
    input.focus();
  };
  dialog.querySelector('form')!.onsubmit = (event) => {
    event.preventDefault();
    let action: EditorLinkAction;
    if (tab === 'document') {
      if (!destination.value) return;
      action = destination.value.startsWith('customShow:')
        ? {
            kind: 'customShow',
            id: Number(destination.value.slice(11)),
            showAndReturn: returnBox.checked,
          }
        : slides.some((slide) => slide.key === destination.value)
          ? { kind: 'slide', slide: destination.value }
          : {
              kind: destination.value as
                | 'nextSlide'
                | 'prevSlide'
                | 'firstSlide'
                | 'lastSlide'
                | 'lastSlideViewed'
                | 'endShow',
            };
    } else {
      const input = field(tab === 'email' ? 'email' : 'address');
      if (!input.value.trim()) {
        input.focus();
        return;
      }
      const originalQuery =
        initial?.action.kind === 'url' && /^mailto:/i.test(initial.action.url)
          ? initial.action.url.split('?')[1]
          : undefined;
      const parameters = new URLSearchParams(originalQuery);
      if (field('subject').value) parameters.set('subject', field('subject').value);
      else parameters.delete('subject');
      const query = [...parameters]
        .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
        .join('&');
      const url =
        tab === 'web'
          ? input.value.trim()
          : !emailEdited && initial?.action.kind === 'url'
            ? initial.action.url
            : `mailto:${encodeURIComponent(field('email').value.trim()).replace(/%40/g, '@')}${query ? '?' + query : ''}`;
      action = { kind: 'url', url };
    }
    if (selectedText !== undefined && !display.value.length) {
      if (selectedText === '')
        display.value =
          tab === 'document'
            ? (destination.selectedOptions[0]?.text ?? '')
            : tab === 'email'
              ? field('email').value.trim()
              : field('address').value.trim();
      if (!display.value.length) {
        display.focus();
        return;
      }
    }
    void save({ action, tooltip });
  };
  update();
  document.body.append(dialog);
  dialog.showModal();
  (tab === 'document' ? destination : field(tab === 'email' ? 'email' : 'address')).focus();
}
