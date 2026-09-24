import type { ParagraphProperties, ParagraphSettings } from '@office-kit/pptx';

type ParagraphValues = Record<string, string>;

export function paragraphDialogValues(initial: Partial<ParagraphProperties>): ParagraphValues {
  const spacing = initial.lineSpacing ?? { kind: 'pct', value: 1 };
  const align =
    (
      {
        l: 'left',
        ctr: 'center',
        r: 'right',
        just: 'justify',
        dist: 'distribute',
        justLow: 'justify',
        thaiDist: 'distribute',
      } as Record<string, string>
    )[initial.align ?? ''] ??
    initial.align ??
    'left';
  return {
    align,
    leftEmu: String((initial.marL ?? 0) / 914400),
    rightEmu: String((initial.marR ?? 0) / 914400),
    special: initial.indent ? (initial.indent < 0 ? 'hanging' : 'first') : 'none',
    indentBy: String(Math.abs(initial.indent ?? 0) / 914400),
    beforePts: String(initial.spcBefPts ?? 0),
    afterPts: String(initial.spcAftPts ?? 0),
    line:
      spacing.kind === 'pts'
        ? 'pts'
        : [1, 1.5, 2].includes(spacing.value)
          ? String(spacing.value)
          : 'pct',
    at: String(spacing.value),
  };
}

/** Blank controls represent differing selected values, not zero. */
export function commonParagraphDialogValues(
  paragraphs: Partial<ParagraphProperties>[],
): ParagraphValues {
  const values = paragraphs.length
    ? paragraphs.map(paragraphDialogValues)
    : [paragraphDialogValues({})];
  return Object.fromEntries(
    Object.entries(values[0]!).map(([key, value]) => [
      key,
      values.every((item) => item[key] === value) ? value : '',
    ]),
  );
}

export function openParagraphDialog(
  initial: Partial<ParagraphProperties>[],
  apply: (settings: ParagraphSettings) => Promise<boolean>,
  restoreFocus: () => void,
) {
  const dialog = document.createElement('dialog');
  dialog.className = 'section-dialog paragraph-dialog';
  dialog.setAttribute('aria-label', 'Paragraph');
  dialog.innerHTML = `<form><h2>Paragraph</h2>
    <label>Alignment<select name="align" aria-label="Alignment"><option value="left">Left</option><option value="center">Centered</option><option value="right">Right</option><option value="justify">Justified</option><option value="distribute">Distributed</option></select></label>
    <fieldset><legend>Indentation</legend>
    <label>Before text<input name="leftEmu" aria-label="Before text" type="number" min="0" max="56" step="any" required> in</label>
    <label>After text<input name="rightEmu" aria-label="After text" type="number" min="0" max="56" step="any" required> in</label>
    <label>Special<select name="special" aria-label="Special"><option value="none">(none)</option><option value="first">First line</option><option value="hanging">Hanging</option></select></label>
    <label>By<input name="indentBy" aria-label="By" type="number" min="0" max="56" step="any" required> in</label></fieldset>
    <fieldset><legend>Spacing</legend>
    <label>Before<input name="beforePts" aria-label="Before" type="number" min="0" max="10000" step="any" required> pt</label>
    <label>After<input name="afterPts" aria-label="After" type="number" min="0" max="10000" step="any" required> pt</label>
    <label>Line spacing<select name="line" aria-label="Line spacing"><option value="1">Single</option><option value="1.5">1.5 Lines</option><option value="2">Double</option><option value="pts">Exactly</option><option value="pct">Multiple</option></select></label>
    <label>At<input name="at" aria-label="At" type="number" min="0" max="10000" step="any" required></label></fieldset>
    <div><button type="button" data-cancel>Cancel</button><button type="submit">OK</button></div></form>`;
  const control = (name: string) =>
    dialog.querySelector<HTMLInputElement | HTMLSelectElement>(`[name="${name}"]`)!;
  const values = commonParagraphDialogValues(initial);
  for (const [name, value] of Object.entries(values)) {
    const input = control(name);
    if (input instanceof HTMLSelectElement) {
      const blank = document.createElement('option');
      blank.value = '';
      blank.textContent = '';
      blank.disabled = true;
      input.prepend(blank);
    }
    input.value = value;
    if (input instanceof HTMLInputElement) input.required = value !== '';
    if (value === '') input.setAttribute('aria-description', 'Mixed values');
  }
  const sync = () => {
    control('indentBy').disabled = !['first', 'hanging'].includes(control('special').value);
    control('at').disabled = !['pts', 'pct'].includes(control('line').value);
  };
  sync();
  const changed = new Set<string>();
  dialog.addEventListener('change', (event) => {
    const target = event.target as HTMLInputElement;
    if (target.name) {
      changed.add(target.name);
      target.removeAttribute('aria-description');
    }
    if (target.name === 'special' && control('indentBy').value === '')
      control('indentBy').value = '0';
    if (target.name === 'line')
      control('at').value =
        target.value === 'pts' ? '18' : target.value === 'pct' ? '1' : target.value;
    sync();
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
    const settings: ParagraphSettings = {};
    const edited = (name: string) => changed.has(name) && control(name).value !== '';
    if (edited('align'))
      settings.align = control('align').value as NonNullable<ParagraphSettings['align']>;
    for (const key of ['leftEmu', 'rightEmu'] as const)
      if (edited(key)) settings[key] = Math.round(Number(control(key).value) * 914400);
    for (const key of ['beforePts', 'afterPts'] as const)
      if (edited(key)) settings[key] = Number(control(key).value);
    if (
      (edited('special') || edited('indentBy')) &&
      control('special').value !== '' &&
      control('indentBy').value !== ''
    )
      settings.firstLineEmu =
        Math.round(Number(control('indentBy').value) * 914400) *
        (control('special').value === 'hanging'
          ? -1
          : control('special').value === 'first'
            ? 1
            : 0);
    if (
      (edited('line') || edited('at')) &&
      control('line').value !== '' &&
      control('at').value !== ''
    ) {
      const line = control('line').value;
      settings.lineSpacing = {
        kind: line === 'pts' ? 'pts' : 'pct',
        value: ['pts', 'pct'].includes(line) ? Number(control('at').value) : Number(line),
      };
    }
    if (!Object.keys(settings).length) {
      close();
      return;
    }
    for (const button of dialog.querySelectorAll('button')) button.disabled = true;
    dialog.close();
    if (await apply(settings)) close();
    else {
      for (const button of dialog.querySelectorAll('button')) button.disabled = false;
      dialog.showModal();
    }
  });
  document.body.append(dialog);
  dialog.showModal();
  control('align').focus();
}
