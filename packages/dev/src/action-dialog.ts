import type { EditorLink, EditorLinkAction, EditorActionSounds } from './editor.ts';

const destinations = [
  ['nextSlide', 'Next Slide'],
  ['prevSlide', 'Previous Slide'],
  ['firstSlide', 'First Slide'],
  ['lastSlide', 'Last Slide'],
  ['lastSlideViewed', 'Last Slide Viewed'],
  ['endShow', 'End Show'],
] as const;

export function openActionDialog(
  initial: { click: EditorLink | null; hover: EditorLink | null; sounds: EditorActionSounds },
  slides: Array<{ key: string; title: string }>,
  apply: (links: {
    click: EditorLink | null;
    hover: EditorLink | null;
    sounds: EditorActionSounds;
  }) => Promise<boolean>,
  restoreFocus: () => void,
  customShows: Array<{ id: number; name: string }> = [],
) {
  const dialog = document.createElement('dialog');
  dialog.className = 'section-dialog action-dialog';
  dialog.setAttribute('aria-label', 'Action Settings');
  dialog.innerHTML = `<form><h2>Action Settings</h2>
    <nav role="tablist" aria-label="Action trigger"><button type="button" role="tab" data-trigger="click">Mouse Click</button><button type="button" role="tab" data-trigger="hover">Mouse Over</button></nav>
    <fieldset><legend>Action on click</legend>
      <label><input type="radio" name="mode" value="none">None</label>
      <label><input type="radio" name="mode" value="link">Hyperlink to:</label>
      <select aria-label="Hyperlink to"></select>
      <button type="button" data-edit-destination>Edit...</button>
    </fieldset><label class="action-sound-label"><input type="checkbox" data-play-sound>Play sound:</label>
    <select aria-label="Sound" data-sound><option value="none">[No Sound]</option><option value="stop">[Stop Previous Sound]</option><option value="other">Other Sound...</option></select>
    <input type="file" accept=".wav,audio/wav,audio/x-wav" data-sound-file hidden>
    <p role="alert"></p>
    <footer><button type="button" data-cancel>Cancel</button><button type="submit">OK</button></footer></form>`;
  const select = dialog.querySelector('select')!;
  for (const [value, label] of destinations) select.add(new Option(label, value));
  select.add(new Option('Slide...', 'slide'));
  select.add(new Option('URL...', 'url'));
  select.add(new Option('Custom Show...', 'customShow'));
  const drafts = { ...initial, sounds: structuredClone(initial.sounds) };
  const soundChoices = structuredClone(initial.sounds);
  const embeddedSounds = { click: initial.sounds.click.sound, hover: initial.sounds.hover.sound };
  const soundEnabled = {
    click: !!initial.sounds.click.sound || initial.sounds.click.stopPrevious,
    hover: !!initial.sounds.hover.sound || initial.sounds.hover.stopPrevious,
  };
  let trigger: 'click' | 'hover' = 'click';
  const captureSound = () => {
    drafts.sounds[trigger] = soundEnabled[trigger]
      ? structuredClone(soundChoices[trigger])
      : { sound: null, stopPrevious: false };
  };
  let action: EditorLinkAction = initial.click?.action ?? { kind: 'nextSlide' };
  select.value = action.kind;
  const none = dialog.querySelector<HTMLInputElement>('[value="none"]')!;
  const link = dialog.querySelector<HTMLInputElement>('[value="link"]')!;
  const edit = dialog.querySelector<HTMLButtonElement>('[data-edit-destination]')!;
  none.checked = !initial.click;
  link.checked = !!initial.click;
  const playSound = dialog.querySelector<HTMLInputElement>('[data-play-sound]')!;
  const soundSelect = dialog.querySelector<HTMLSelectElement>('[data-sound]')!;
  const soundFile = dialog.querySelector<HTMLInputElement>('[data-sound-file]')!;
  let loadingSound = false;
  const renderSound = () => {
    soundSelect.querySelector('[value="embedded"]')?.remove();
    const value = soundChoices[trigger];
    const embedded = embeddedSounds[trigger];
    if (embedded)
      soundSelect.add(
        new Option(embedded.name || 'Sound', 'embedded'),
        soundSelect.options.length - 1,
      );
    playSound.checked = soundEnabled[trigger];
    soundSelect.value = value.sound ? 'embedded' : value.stopPrevious ? 'stop' : 'none';
  };
  let saving = false;
  const update = () => {
    select.disabled = none.checked || saving;
    soundSelect.disabled = !playSound.checked || saving || loadingSound;
    edit.disabled =
      none.checked || saving || !['url', 'slide', 'customShow'].includes(select.value);
  };
  const capture = () => {
    drafts[trigger] = none.checked ? null : { action, tooltip: drafts[trigger]?.tooltip ?? null };
  };
  const tabs = Array.from(dialog.querySelectorAll<HTMLButtonElement>('[data-trigger]'));
  const showTrigger = (next: 'click' | 'hover') => {
    capture();
    trigger = next;
    renderSound();
    action = drafts[trigger]?.action ?? { kind: 'nextSlide' };
    none.checked = !drafts[trigger];
    link.checked = !!drafts[trigger];
    select.value = action.kind;
    dialog.querySelector('legend')!.textContent =
      trigger === 'click' ? 'Action on click' : 'Action on mouse over';
    tabs.forEach((tab) => {
      const active = tab.dataset.trigger === trigger;
      tab.setAttribute('aria-selected', String(active));
      tab.tabIndex = active ? 0 : -1;
    });
    update();
  };
  tabs.forEach((tab) => {
    tab.onclick = () => showTrigger(tab.dataset.trigger as 'click' | 'hover');
    tab.onkeydown = (event) => {
      if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
      event.preventDefault();
      const next =
        event.key === 'Home'
          ? 'click'
          : event.key === 'End'
            ? 'hover'
            : trigger === 'click'
              ? 'hover'
              : 'click';
      showTrigger(next);
      tabs.find((item) => item.dataset.trigger === next)!.focus();
    };
  });
  playSound.onchange = () => {
    soundEnabled[trigger] = playSound.checked;
    captureSound();
    update();
  };
  soundSelect.onchange = () => {
    if (soundSelect.value === 'other') {
      soundFile.value = '';
      soundFile.click();
      return;
    }
    soundChoices[trigger] =
      soundSelect.value === 'embedded'
        ? { sound: embeddedSounds[trigger], stopPrevious: false }
        : { sound: null, stopPrevious: soundSelect.value === 'stop' };
    captureSound();
  };
  soundFile.addEventListener('cancel', () => {
    renderSound();
    update();
  });
  soundFile.onchange = async () => {
    const file = soundFile.files?.[0];
    if (!file) {
      renderSound();
      update();
      return;
    }
    const target = trigger;
    loadingSound = true;
    update();
    try {
      if (file.size > 20_000_000) throw new Error('Choose a WAV sound under 20 MB.');
      const bytes = new Uint8Array(await file.arrayBuffer());
      const signature = new TextDecoder();
      if (
        signature.decode(bytes.slice(0, 4)) !== 'RIFF' ||
        signature.decode(bytes.slice(8, 12)) !== 'WAVE'
      )
        throw new Error('Choose a WAV sound.');
      let binary = '';
      for (let offset = 0; offset < bytes.length; offset += 8192)
        binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
      soundChoices[target] = {
        sound: { name: file.name, base64: btoa(binary) },
        stopPrevious: false,
      };
      embeddedSounds[target] = soundChoices[target].sound;
      soundEnabled[target] = true;
      drafts.sounds[target] = structuredClone(soundChoices[target]);
      dialog.querySelector('[role="alert"]')!.textContent = '';
    } catch (error) {
      dialog.querySelector('[role="alert"]')!.textContent =
        error instanceof Error ? error.message : 'Unable to read sound.';
    }
    loadingSound = false;
    renderSound();
    update();
  };
  showTrigger('click');
  none.onchange = link.onchange = update;
  const close = () => {
    dialog.close();
    dialog.remove();
    restoreFocus();
  };
  dialog.addEventListener('cancel', (event) => {
    event.preventDefault();
    if (!saving) close();
  });
  dialog.querySelector<HTMLButtonElement>('[data-cancel]')!.onclick = close;
  const editDestination = () => {
    const kind = select.value;
    const child = document.createElement('dialog');
    child.className = 'section-dialog action-dialog';
    const title =
      kind === 'url'
        ? 'Hyperlink to URL'
        : kind === 'customShow'
          ? 'Hyperlink to Custom Show'
          : 'Hyperlink to Slide';
    child.setAttribute('aria-label', title);
    child.innerHTML = `<form><h2>${title}</h2><label>${kind === 'url' ? 'URL:' : kind === 'customShow' ? 'Custom show:' : 'Slide:'}${kind === 'url' ? '<input required>' : '<select size="7" required></select>'}</label><footer><button type="button">Cancel</button><button type="submit">OK</button></footer></form>`;
    if (kind === 'customShow')
      child
        .querySelector('footer')!
        .insertAdjacentHTML(
          'beforebegin',
          '<label><input type="checkbox" data-return>Show and return</label>',
        );
    const returnBox = child.querySelector<HTMLInputElement>('[data-return]');
    if (returnBox) returnBox.checked = action.kind === 'customShow' && action.showAndReturn;
    const field = child.querySelector<HTMLInputElement | HTMLSelectElement>('input,select')!;
    if (field instanceof HTMLSelectElement) {
      if (kind === 'customShow') {
        customShows.forEach((show) => field.add(new Option(show.name, String(show.id))));
        if (action.kind === 'customShow') field.value = String(action.id);
      } else
        slides.forEach((slide, index) =>
          field.add(
            new Option(`${index + 1}. ${slide.title || 'Slide ' + (index + 1)}`, slide.key),
          ),
        );
      if (action.kind === 'slide') field.value = action.slide;
    } else if (action.kind === 'url') field.value = action.url;
    const closeChild = () => {
      child.close();
      child.remove();
      select.value = action.kind;
      update();
      select.focus();
    };
    child.addEventListener('cancel', (event) => {
      event.preventDefault();
      closeChild();
    });
    child.querySelector('button')!.onclick = closeChild;
    child.querySelector('form')!.onsubmit = (event) => {
      event.preventDefault();
      if (!field.value.trim()) return;
      action =
        kind === 'url'
          ? { kind: 'url', url: field.value.trim() }
          : kind === 'customShow'
            ? { kind: 'customShow', id: Number(field.value), showAndReturn: returnBox!.checked }
            : { kind: 'slide', slide: field.value };
      closeChild();
    };
    document.body.append(child);
    child.showModal();
    field.focus();
  };
  edit.onclick = editDestination;
  select.onchange = () => {
    if (['url', 'slide', 'customShow'].includes(select.value)) editDestination();
    else action = { kind: select.value as (typeof destinations)[number][0] };
    update();
  };
  dialog.querySelector('form')!.onsubmit = async (event) => {
    event.preventDefault();
    if (saving || loadingSound) return;
    capture();
    saving = true;
    dialog
      .querySelectorAll<HTMLButtonElement | HTMLInputElement>('button,input')
      .forEach((element) => {
        element.disabled = true;
      });
    update();
    dialog.close();
    try {
      if (await apply(drafts)) {
        close();
        return;
      }
      dialog.querySelector('[role="alert"]')!.textContent =
        'Unable to save action. Check the current selection and try again.';
    } catch (error) {
      dialog.querySelector('[role="alert"]')!.textContent =
        error instanceof Error ? error.message : 'Unable to save action.';
    }
    saving = false;
    dialog
      .querySelectorAll<HTMLButtonElement | HTMLInputElement>('button,input')
      .forEach((element) => {
        element.disabled = false;
      });
    update();
    dialog.showModal();
  };
  document.body.append(dialog);
  update();
  dialog.showModal();
  (initial.click ? link : none).focus();
}
