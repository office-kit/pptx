import type { EditorShape } from './editor.ts';

function fontValues(initial: EditorShape['format']): Record<string, string | boolean> {
  return {
    font: initial.font || 'Calibri',
    size: String(initial.size ?? 18),
    style: initial.bold
      ? initial.italic
        ? 'both'
        : 'bold'
      : initial.italic
        ? 'italic'
        : 'regular',
    color: /^#?[0-9a-f]{6}$/i.test(initial.color ?? '')
      ? '#' + initial.color!.replace(/^#/, '').toLowerCase()
      : '#000000',
    underline: ['sng', 'dbl'].includes(String(initial.underline))
      ? String(initial.underline)
      : 'none',
    strike: !!initial.strike && initial.strike !== 'noStrike',
    cap: initial.cap ?? 'none',
    baseline: String(initial.baseline ? (initial.baseline > 0 ? 0.3 : -0.25) : 0),
    offset: String(Math.abs(initial.baseline ?? 0.3) * 100),
    spacing: initial.spc ? (initial.spc > 0 ? 'expanded' : 'condensed') : 'normal',
    spacingBy: String(Math.abs(initial.spc ?? 100) / 100),
    kerning: (initial.kern ?? 0) > 0,
    kern: String((initial.kern || 1200) / 100),
  };
}

// Return only changed properties so opening this dialog on mixed text does not
// flatten unrelated formatting when one control is edited.
export function openFontDialog(
  initial: EditorShape['format'][],
  apply: (format: EditorShape['format']) => Promise<boolean>,
  restoreFocus: () => void,
  initialTab: 'font' | 'spacing' = 'font',
) {
  const dialog = document.createElement('dialog');
  dialog.className = 'section-dialog font-dialog';
  dialog.setAttribute('aria-label', 'Font');
  dialog.innerHTML = `<form><h2>Font</h2>
    <div role="tablist" aria-label="Font settings"><button type="button" role="tab" id="font-tab" aria-controls="font-panel" data-tab="font">Font</button><button type="button" role="tab" id="spacing-tab" aria-controls="spacing-panel" data-tab="spacing">Character Spacing</button></div>
    <section role="tabpanel" id="font-panel" aria-labelledby="font-tab">
    <label>Font<input name="font" aria-label="Font" required></label>
    <label>Font style<select name="style" aria-label="Font style"><option value="regular">Regular</option><option value="bold">Bold</option><option value="italic">Italic</option><option value="both">Bold Italic</option></select></label>
    <label>Size<input name="size" aria-label="Size" type="number" min="1" max="400" step="0.1" required></label>
    <label>Font color<input name="color" aria-label="Font color" type="color"></label>
    <label>Underline style<select name="underline" aria-label="Underline style"><option value="none">None</option><option value="sng">Single</option><option value="dbl">Double</option></select></label>
    <fieldset><legend>Effects</legend>
      <label class="font-effect"><input name="strike" type="checkbox">Strikethrough</label>
      <label>Capitalization<select name="cap" aria-label="Capitalization"><option value="none">None</option><option value="small">Small caps</option><option value="all">All caps</option></select></label>
      <label>Position<select name="baseline" aria-label="Position"><option value="0">Normal</option><option value="0.3">Superscript</option><option value="-0.25">Subscript</option></select></label>
      <label>Offset<input name="offset" aria-label="Offset" type="number" min="0" max="100" step="0.1" required> %</label>
    </fieldset></section>
    <section role="tabpanel" id="spacing-panel" aria-labelledby="spacing-tab">
      <label>Spacing<select name="spacing" aria-label="Spacing"><option value="normal">Normal</option><option value="expanded">Expanded</option><option value="condensed">Condensed</option></select></label>
      <label>By<input name="spacingBy" aria-label="By" type="number" min="0" max="20" step="0.1" required> pt</label>
      <label class="font-effect"><input name="kerning" type="checkbox">Kerning for fonts</label>
      <label>Points and above<input name="kern" aria-label="Points and above" type="number" min="0" max="4000" step="0.1" required></label>
    </section>
    <div><button type="button" data-cancel>Cancel</button><button type="submit">OK</button></div></form>`;
  const control = (name: string) =>
    dialog.querySelector<HTMLInputElement | HTMLSelectElement>(`[name="${name}"]`)!;
  const values = (initial.length ? initial : [{}]).map(fontValues);
  for (const [name, value] of Object.entries(values[0]!)) {
    const input = control(name);
    const mixed = values.some((item) => item[name] !== value);
    if (input instanceof HTMLInputElement && input.type === 'checkbox') {
      input.checked = value === true;
      input.indeterminate = mixed;
    } else {
      if (input instanceof HTMLSelectElement) {
        const blank = document.createElement('option');
        blank.value = '';
        blank.disabled = true;
        input.prepend(blank);
      }
      input.value = mixed ? '' : String(value);
      if (input instanceof HTMLInputElement && mixed) input.required = false;
    }
    if (mixed) {
      input.setAttribute('aria-description', 'Mixed values');
      input.classList.add('mixed-value');
    }
  }
  const syncControls = () => {
    control('offset').disabled = !['0.3', '-0.25'].includes(control('baseline').value);
    control('spacingBy').disabled = !['expanded', 'condensed'].includes(control('spacing').value);
    control('kern').disabled =
      !(control('kerning') as HTMLInputElement).checked ||
      (control('kerning') as HTMLInputElement).indeterminate;
  };
  syncControls();
  const tabs = [...dialog.querySelectorAll<HTMLButtonElement>('[data-tab]')];
  const selectTab = (name: string) => {
    for (const tab of tabs) {
      const selected = tab.dataset.tab === name;
      tab.setAttribute('aria-selected', String(selected));
      tab.tabIndex = selected ? 0 : -1;
      const panel = dialog.querySelector<HTMLElement>(`#${tab.getAttribute('aria-controls')}`)!;
      panel.hidden = !selected;
    }
  };
  for (const tab of tabs) {
    tab.addEventListener('click', () => selectTab(tab.dataset.tab!));
    tab.addEventListener('keydown', (event) => {
      if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
      event.preventDefault();
      const next =
        event.key === 'Home'
          ? tabs[0]!
          : event.key === 'End'
            ? tabs[1]!
            : tabs.find((item) => item !== tab)!;
      selectTab(next.dataset.tab!);
      next.focus();
    });
  }
  selectTab(initialTab);
  const changed = new Set<string>();
  dialog.addEventListener('change', (event) => {
    const target = event.target as HTMLInputElement;
    if (target.name) changed.add(target.name);
    target.removeAttribute('aria-description');
    target.classList.remove('mixed-value');
    if (target.name === 'spacing' && !control('spacingBy').value) control('spacingBy').value = '1';
    if (target.name === 'kerning' && !control('kern').value) control('kern').value = '12';
    if (target.name === 'baseline')
      control('offset').value = target.value === '-0.25' ? '25' : '30';
    syncControls();
  });
  const close = () => {
    dialog.close();
    dialog.remove();
    restoreFocus();
  };
  dialog.querySelector('[data-cancel]')!.addEventListener('click', close);
  dialog.addEventListener('cancel', (event) => {
    event.preventDefault();
    close();
  });
  dialog.querySelector('form')!.addEventListener('submit', async (event) => {
    event.preventDefault();
    const format: EditorShape['format'] = {};
    const edited = (name: string) => changed.has(name) && control(name).value !== '';
    if (edited('font')) format.font = control('font').value.trim();
    if (edited('size')) format.size = Number(control('size').value);
    if (edited('color')) format.color = control('color').value;
    if (edited('style')) {
      const style = control('style').value;
      format.bold = style === 'bold' || style === 'both';
      format.italic = style === 'italic' || style === 'both';
    }
    if (edited('underline')) format.underline = control('underline').value;
    if (changed.has('strike')) format.strike = (control('strike') as HTMLInputElement).checked;
    if (edited('cap')) format.cap = control('cap').value as 'none' | 'small' | 'all';
    if (
      (edited('baseline') || edited('offset')) &&
      control('baseline').value !== '' &&
      (control('baseline').value === '0' || control('offset').value !== '')
    ) {
      format.baseline =
        (Math.sign(Number(control('baseline').value)) * Number(control('offset').value)) / 100;
    }
    if (
      (edited('spacing') || edited('spacingBy')) &&
      control('spacing').value !== '' &&
      (control('spacing').value === 'normal' || control('spacingBy').value !== '')
    ) {
      const direction = control('spacing').value;
      format.spc =
        direction === 'normal'
          ? 0
          : Math.round(Number(control('spacingBy').value) * 100) *
            (direction === 'condensed' ? -1 : 1);
    }
    if (changed.has('kerning') || edited('kern')) {
      format.kern = (control('kerning') as HTMLInputElement).checked
        ? Math.round(Number(control('kern').value) * 100)
        : 0;
    }
    if (!Object.keys(format).length) {
      close();
      return;
    }
    for (const button of dialog.querySelectorAll('button')) button.disabled = true;
    // Release modal inertness before the editor rebuilds and reselects its text.
    dialog.close();
    if (await apply(format)) close();
    else {
      for (const button of dialog.querySelectorAll('button')) button.disabled = false;
      dialog.showModal();
    }
  });
  document.body.append(dialog);
  dialog.showModal();
  control(initialTab === 'font' ? 'font' : 'spacing').focus();
}
