import type { EditorModel } from './editor.ts';

type Shows = EditorModel['customShows'];
export function openCustomShows(
  initial: Shows,
  slides: Array<{ key: string; title: string }>,
  apply: (shows: Shows) => Promise<boolean>,
  start: (id: number) => void,
  restoreFocus: () => void,
) {
  let shows = structuredClone(initial);
  let saving = false;
  const dialog = document.createElement('dialog');
  dialog.className = 'section-dialog custom-shows-dialog';
  dialog.setAttribute('aria-label', 'Custom Shows');
  dialog.innerHTML = `<h2>Custom Shows</h2><select size="8" aria-label="Custom shows" style="width:100%;min-width:300px"></select>
    <div><button data-add aria-label="New custom show">+</button><button data-remove aria-label="Delete custom show">−</button><button data-edit-show>Edit...</button><button data-copy>Copy</button></div>
    <p role="alert"></p><footer><button data-start>Start Show</button><button data-close>Close</button></footer>`;
  const list = dialog.querySelector('select')!;
  const button = (selector: string) => dialog.querySelector<HTMLButtonElement>(selector)!;
  const render = (id?: number) => {
    list.replaceChildren(...shows.map((show) => new Option(show.name, String(show.id))));
    if (id !== undefined) list.value = String(id);
    if (list.selectedIndex < 0 && shows.length) list.selectedIndex = 0;
    update();
  };
  const chosen = () => shows.find((show) => String(show.id) === list.value);
  const update = () => {
    list.disabled = saving;
    for (const control of dialog.querySelectorAll<HTMLButtonElement>('button'))
      control.disabled = saving;
    for (const selector of ['[data-remove]', '[data-edit-show]', '[data-copy]'])
      button(selector).disabled = saving || !chosen();
    button('[data-start]').disabled = saving || !chosen()?.slides.length;
  };
  const commit = async (next: Shows, id?: number) => {
    saving = true;
    update();
    dialog.querySelector('[role="alert"]')!.textContent = '';
    try {
      if (!(await apply(next))) return false;
      shows = structuredClone(next);
      render(id);
      return true;
    } catch (error) {
      dialog.querySelector('[role="alert"]')!.textContent =
        error instanceof Error ? error.message : String(error);
      return false;
    } finally {
      saving = false;
      update();
    }
  };
  const nextId = () => {
    let id = 0;
    const used = new Set(shows.map((show) => show.id));
    while (used.has(id)) id++;
    return id;
  };
  const uniqueName = (base: string) => {
    let name = base;
    for (let n = 2; shows.some((show) => show.name === name); n++) name = `${base} ${n}`;
    return name;
  };
  const define = (source?: Shows[number]) => {
    const show = structuredClone(
      source ?? { id: nextId(), name: uniqueName('Custom Show 1'), slides: [] },
    );
    const edit = document.createElement('dialog');
    edit.className = 'section-dialog custom-show-definition';
    edit.setAttribute('aria-label', 'Define Custom Show');
    edit.innerHTML = `<form><h2>Define Custom Show</h2><label>Slide show name: <input data-name required></label>
      <div style="display:flex;gap:12px;margin-top:16px;align-items:center">
      <label>Slides in presentation:<br><select data-source multiple size="10" style="min-width:200px"></select></label>
      <button type="button" data-include>Add →</button>
      <label>Slides in custom show:<br><select data-sequence size="10" style="min-width:200px"></select></label>
      <div style="display:grid;gap:8px"><button type="button" data-up aria-label="Move slide up">↑</button><button type="button" data-down aria-label="Move slide down">↓</button><button type="button" data-exclude>Remove</button></div></div>
      <p role="alert"></p><footer><button type="button" data-cancel>Cancel</button><button type="submit">OK</button></footer></form>`;
    const name = edit.querySelector<HTMLInputElement>('[data-name]')!;
    const available = edit.querySelector<HTMLSelectElement>('[data-source]')!;
    const sequence = edit.querySelector<HTMLSelectElement>('[data-sequence]')!;
    const control = (selector: string) => edit.querySelector<HTMLButtonElement>(selector)!;
    const label = (key: string) => {
      const index = slides.findIndex((slide) => slide.key === key);
      return `${index + 1}. ${slides[index]?.title || `Slide ${index + 1}`}`;
    };
    name.value = show.name;
    available.replaceChildren(...slides.map((slide) => new Option(label(slide.key), slide.key)));
    const updateDefinition = () => {
      const position = sequence.selectedIndex;
      control('[data-include]').disabled = !available.selectedOptions.length;
      control('[data-up]').disabled = position <= 0;
      control('[data-down]').disabled = position < 0 || position >= show.slides.length - 1;
      control('[data-exclude]').disabled = position < 0;
    };
    const renderSequence = (position: number) => {
      sequence.replaceChildren(...show.slides.map((key, i) => new Option(label(key), String(i))));
      sequence.selectedIndex = Math.min(position, show.slides.length - 1);
      updateDefinition();
    };
    available.onchange = sequence.onchange = updateDefinition;
    control('[data-include]').onclick = () => {
      show.slides.push(...Array.from(available.selectedOptions, (option) => option.value));
      renderSequence(show.slides.length - 1);
    };
    control('[data-exclude]').onclick = () => {
      const position = sequence.selectedIndex;
      if (position < 0) return;
      show.slides.splice(position, 1);
      renderSequence(position);
    };
    for (const [selector, direction] of [
      ['[data-up]', -1],
      ['[data-down]', 1],
    ] as const)
      control(selector).onclick = () => {
        const position = sequence.selectedIndex;
        const target = position + direction;
        if (position < 0 || target < 0 || target >= show.slides.length) return;
        [show.slides[position], show.slides[target]] = [
          show.slides[target]!,
          show.slides[position]!,
        ];
        renderSequence(target);
      };
    let submitting = false;
    control('[data-cancel]').onclick = () => edit.close();
    edit.oncancel = (event) => {
      if (submitting) event.preventDefault();
    };
    edit.onclose = () => {
      edit.remove();
      list.focus();
    };
    edit.querySelector('form')!.onsubmit = async (event) => {
      event.preventDefault();
      if (submitting) return;
      if (shows.some((other) => other.id !== show.id && other.name === name.value)) {
        edit.querySelector('[role="alert"]')!.textContent =
          'A custom show with this name already exists.';
        return;
      }
      submitting = true;
      for (const element of edit.querySelectorAll<
        HTMLInputElement | HTMLButtonElement | HTMLSelectElement
      >('input,button,select'))
        element.disabled = true;
      show.name = name.value;
      const next = source
        ? shows.map((other) => (other.id === source.id ? show : other))
        : [...shows, show];
      if (await commit(next, show.id)) edit.close();
      else {
        edit.querySelector('[role="alert"]')!.textContent =
          dialog.querySelector('[role="alert"]')!.textContent ||
          'The custom show could not be saved.';
        for (const element of edit.querySelectorAll<
          HTMLInputElement | HTMLButtonElement | HTMLSelectElement
        >('input,button,select'))
          element.disabled = false;
        updateDefinition();
      }
      submitting = false;
    };
    renderSequence(0);
    document.body.append(edit);
    edit.showModal();
    name.focus();
    name.select();
  };
  button('[data-add]').onclick = () => define();
  button('[data-edit-show]').onclick = () => {
    const show = chosen();
    if (show) define(show);
  };
  list.ondblclick = () => {
    if (!saving && chosen()) define(chosen());
  };
  list.onchange = update;
  button('[data-remove]').onclick = () => void commit(shows.filter((show) => show !== chosen()));
  button('[data-copy]').onclick = () => {
    const show = chosen();
    if (!show) return;
    const copy = { ...structuredClone(show), id: nextId(), name: uniqueName(`${show.name} Copy`) };
    void commit([...shows, copy], copy.id);
  };
  button('[data-close]').onclick = () => dialog.close();
  button('[data-start]').onclick = () => {
    const show = chosen();
    if (show?.slides.length) {
      dialog.close();
      start(show.id);
    }
  };
  dialog.oncancel = (event) => {
    if (saving) event.preventDefault();
  };
  dialog.onclose = () => {
    dialog.remove();
    if (!document.body.classList.contains('presenting')) restoreFocus();
  };
  render();
  document.body.append(dialog);
  dialog.showModal();
  list.focus();
}
