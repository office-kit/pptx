import type { EditorModel } from './editor.ts';

type Settings = EditorModel['showProperties'];
export function openShowProperties(
  initial: Settings,
  slideCount: number,
  shows: EditorModel['customShows'],
  apply: (settings: Settings) => Promise<boolean>,
  restoreFocus: () => void,
) {
  const dialog = document.createElement('dialog');
  dialog.className = 'section-dialog show-properties-dialog';
  dialog.setAttribute('aria-label', 'Set Up Show');
  dialog.innerHTML = `<form><h2>Set Up Show</h2>
    <fieldset><legend>Show type</legend>
    <label><input type="radio" name="mode" value="present"> Presented by a speaker (full screen)</label>
    <label><input type="radio" name="mode" value="browse"> Browsed by an individual (window)</label>
    <label class="show-scrollbar-option"><input type="checkbox" data-scrollbar> Show scrollbar</label>
    <label><input type="radio" name="mode" value="kiosk"> Browsed at a kiosk (full screen)</label></fieldset>
    <div class="show-settings-columns"><fieldset><legend>Show options</legend>
    <label><input type="checkbox" data-loop> Loop continuously until 'Esc'</label>
    <label><input type="checkbox" data-no-narration> Show without narration</label>
    <label><input type="checkbox" data-no-animation> Show without animation</label></fieldset>
    <fieldset><legend>Show slides</legend>
    <label><input type="radio" name="slides" value="all"> All</label>
    <div class="show-range"><label><input type="radio" name="slides" value="range"> From:</label> <input type="number" aria-label="From slide" data-start min="1" step="1" required> <label>To: <input type="number" aria-label="To slide" data-end min="1" step="1" required></label></div>
    <label><input type="radio" name="slides" value="customShow"> Custom show:</label><select aria-label="Custom show"></select></fieldset></div>
    <fieldset><legend>Advance slides</legend>
    <label><input type="radio" name="advance" value="manual"> Manually</label>
    <label><input type="radio" name="advance" value="timings"> Using timings, if present</label></fieldset>
    <p role="alert"></p><footer><button type="button" data-cancel>Cancel</button><button type="submit">OK</button></footer></form>`;
  const input = (selector: string) => dialog.querySelector<HTMLInputElement>(selector)!;
  const selected = (name: string) => input(`input[name="${name}"]:checked`).value;
  const select = (name: string, value: string) => {
    input(`input[name="${name}"][value="${value}"]`).checked = true;
  };
  const list = dialog.querySelector('select')!;
  list.replaceChildren(...shows.map((show) => new Option(show.name, String(show.id))));
  if (initial.slides.kind === 'customShow') {
    const id = initial.slides.id;
    if (!shows.some((show) => show.id === id))
      list.add(new Option('Unavailable custom show', String(initial.slides.id)));
    list.value = String(initial.slides.id);
  }
  select('mode', initial.mode.kind);
  select('slides', initial.slides.kind);
  select('advance', initial.useTimings ? 'timings' : 'manual');
  input('[data-scrollbar]').checked =
    initial.mode.kind === 'browse' ? initial.mode.showScrollbar : true;
  input('[data-loop]').checked = initial.loop;
  input('[data-no-narration]').checked = !initial.showNarration;
  input('[data-no-animation]').checked = !initial.showAnimation;
  input('[data-start]').value = String(initial.slides.kind === 'range' ? initial.slides.start : 1);
  input('[data-end]').value = String(
    initial.slides.kind === 'range' ? initial.slides.end : Math.max(1, slideCount),
  );
  for (const selector of ['[data-start]', '[data-end]']) input(selector).max = String(slideCount);
  let saving = false;
  const update = () => {
    for (const control of dialog.querySelectorAll<
      HTMLInputElement | HTMLButtonElement | HTMLSelectElement
    >('input,button,select'))
      control.disabled = saving;
    input('[data-scrollbar]').disabled = saving || selected('mode') !== 'browse';
    const kiosk = selected('mode') === 'kiosk';
    if (kiosk) input('[data-loop]').checked = true;
    input('[data-loop]').disabled = saving || kiosk;
    for (const selector of ['[data-start]', '[data-end]'])
      input(selector).disabled = saving || selected('slides') !== 'range';
    input('input[name="slides"][value="customShow"]').disabled = saving || !list.options.length;
    list.disabled = saving || selected('slides') !== 'customShow';
  };
  dialog.oninput = () => input('[data-end]').setCustomValidity('');
  dialog.onchange = () => {
    input('[data-end]').setCustomValidity('');
    update();
  };
  dialog.querySelector('[data-cancel]')!.addEventListener('click', () => dialog.close());
  dialog.oncancel = (event) => {
    if (saving) event.preventDefault();
  };
  dialog.onclose = () => {
    dialog.remove();
    restoreFocus();
  };
  dialog.querySelector('form')!.onsubmit = async (event) => {
    event.preventDefault();
    if (saving) return;
    const kind = selected('slides');
    const start = input('[data-start]').valueAsNumber,
      end = input('[data-end]').valueAsNumber;
    if (kind === 'range' && start > end) {
      input('[data-end]').setCustomValidity('The last slide must be after the first slide.');
      input('[data-end]').reportValidity();
      return;
    }
    const mode = selected('mode');
    const settings: Settings = {
      mode:
        mode === 'browse'
          ? {
              kind: 'browse',
              showScrollbar: input('[data-scrollbar]').checked,
            }
          : mode === 'kiosk'
            ? {
                kind: 'kiosk',
                restart: initial.mode.kind === 'kiosk' ? initial.mode.restart : 300000,
              }
            : { kind: 'present' },
      slides:
        kind === 'range'
          ? { kind: 'range', start, end }
          : kind === 'customShow'
            ? { kind: 'customShow', id: Number(list.value) }
            : { kind: 'all' },
      loop: input('[data-loop]').checked,
      showNarration: !input('[data-no-narration]').checked,
      showAnimation: !input('[data-no-animation]').checked,
      useTimings: selected('advance') === 'timings',
    };
    saving = true;
    update();
    dialog.querySelector('[role="alert"]')!.textContent = '';
    try {
      if (await apply(settings)) dialog.close();
    } catch (error) {
      dialog.querySelector('[role="alert"]')!.textContent =
        error instanceof Error ? error.message : String(error);
    } finally {
      saving = false;
      update();
    }
  };
  update();
  document.body.append(dialog);
  dialog.showModal();
}
