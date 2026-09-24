import type { EditorShape, EditCommand } from './editor.ts';
import { createShapePaint } from './shape-paint.ts';

/** Native Size & Properties controls. Values displayed in centimeters. */
export function createShapeProperties(options: {
  selection(): EditorShape[];
  disabled(): boolean;
  update(command: Partial<EditCommand>): Promise<unknown>;
  focusCanvas(): void;
}) {
  const pane = document.createElement('aside');
  pane.id = 'shape-properties';
  pane.hidden = true;
  pane.setAttribute('aria-label', 'Format Shape');
  pane.innerHTML = `<header><h2>Format Shape</h2><button aria-label="Close Format Shape">×</button></header>
    <div class="properties-category" role="tablist" aria-label="Shape Options"><button role="tab" aria-selected="false" aria-controls="shape-paint-panel" id="shape-paint-tab">Fill &amp; Line</button><button role="tab" aria-selected="true" aria-controls="shape-size-panel" id="shape-size-tab">Size &amp; Properties</button><button role="tab" aria-selected="false" aria-controls="picture-panel" id="picture-tab" hidden>Picture</button></div>
    <div id="picture-panel" role="tabpanel" aria-labelledby="picture-tab" hidden><details open data-picture-corrections><summary>Picture Corrections</summary><div class="properties-fields"></div></details><details open><summary>Picture Transparency</summary><div class="properties-fields"><label>Transparency<input type="range" min="0" max="100" step="1" aria-label="Picture transparency slider"></label><label>Transparency <span class="properties-value"><input type="number" min="0" max="100" step="1" aria-label="Picture transparency">%</span></label></div></details></div>
    <div id="shape-paint-panel" role="tabpanel" aria-labelledby="shape-paint-tab" hidden></div>
    <div id="shape-size-panel" role="tabpanel" aria-labelledby="shape-size-tab">
    <details open><summary>Size</summary><div class="properties-fields" data-section="size"></div></details>
    <details open><summary>Position</summary><div class="properties-fields" data-section="position"></div></details>
    <details><summary>Text Box</summary><div class="properties-fields" data-section="text"></div></details></div>`;
  document.querySelector('main')!.append(pane);
  const paint = createShapePaint(pane.querySelector('#shape-paint-panel')!, (changes) =>
    options.update({ changes }),
  );
  const tabs = Array.from(pane.querySelectorAll<HTMLButtonElement>('[role="tab"]'));
  function selectTab(tab: HTMLButtonElement) {
    for (const item of tabs) {
      const selected = item === tab;
      item.setAttribute('aria-selected', String(selected));
      item.tabIndex = selected ? 0 : -1;
      pane.querySelector<HTMLElement>(`#${item.getAttribute('aria-controls')}`)!.hidden = !selected;
    }
  }
  for (const tab of tabs) {
    tab.tabIndex = tab.getAttribute('aria-selected') === 'true' ? 0 : -1;
    tab.onclick = () => selectTab(tab);
    tab.onkeydown = (event) => {
      if (['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) {
        event.preventDefault();
        const available = tabs.filter((item) => !item.hidden);
        const next =
          event.key === 'Home'
            ? available[0]!
            : event.key === 'End'
              ? available[available.length - 1]!
              : available[
                  (available.indexOf(tab) +
                    (event.key === 'ArrowLeft' ? -1 : 1) +
                    available.length) %
                    available.length
                ]!;
        selectTab(next);
        next.focus();
      }
    };
  }
  const transparency = Array.from(pane.querySelectorAll<HTMLInputElement>('#picture-panel input'));
  const corrections: Array<{ input: HTMLInputElement; key: 'imageBrightness' | 'imageContrast' }> =
    [];
  for (const [key, name] of [
    ['imageBrightness', 'Brightness'],
    ['imageContrast', 'Contrast'],
  ] as const) {
    for (const type of ['range', 'number']) {
      const label = document.createElement('label');
      label.textContent = name;
      const input = document.createElement('input');
      input.type = type;
      input.min = '-100';
      input.max = '100';
      input.step = '1';
      input.setAttribute(
        'aria-label',
        `Picture ${name.toLowerCase()}${type === 'range' ? ' slider' : ''}`,
      );
      if (type === 'number') {
        const unit = document.createElement('span');
        unit.className = 'properties-value';
        unit.append(input, '%');
        label.append(unit);
      } else label.append(input);
      pane.querySelector('[data-picture-corrections] .properties-fields')!.append(label);
      corrections.push({ input, key });
      input.onchange = () => {
        if (input.disabled) return;
        const value = input.valueAsNumber;
        if (!Number.isFinite(value) || value < -100 || value > 100) {
          input.value = '';
          render();
          return;
        }
        void options.update({ changes: { [key]: value / 100 } });
      };
    }
  }
  for (const input of transparency) {
    input.onchange = () => {
      if (input.disabled) return;
      const value = input.valueAsNumber;
      if (!Number.isFinite(value) || value < 0 || value > 100) {
        input.value = '';
        render();
        return;
      }
      void options.update({ changes: { imageOpacity: 1 - value / 100 } });
    };
  }
  const fields = new Map<string, HTMLInputElement>();
  const definitions = [
    ['h', 'Height', 'size'],
    ['w', 'Width', 'size'],
    ['rotation', 'Rotation', 'size'],
    ['x', 'Horizontal position', 'position'],
    ['y', 'Vertical position', 'position'],
  ] as const;
  for (const [key, name, section] of definitions) {
    const label = document.createElement('label');
    label.textContent = name;
    const input = document.createElement('input');
    input.type = 'number';
    input.setAttribute('aria-label', name);
    input.step = key === 'rotation' ? '1' : '0.01';
    input.min = key === 'rotation' ? '-3600' : key === 'w' || key === 'h' ? '0' : '-5963.92';
    input.max = key === 'rotation' ? '3600' : '5963.92';
    const unit = document.createElement('span');
    unit.className = 'properties-value';
    unit.append(input, key === 'rotation' ? '°' : 'cm');
    label.append(unit);
    pane.querySelector(`[data-section="${section}"]`)!.append(label);
    fields.set(key, input);
    input.onkeydown = (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        input.blur();
      }
      if (event.key === 'Escape') {
        event.preventDefault();
        input.value = '';
        input.blur();
        render();
      }
    };
    input.onchange = () => {
      if (input.value === '' || !input.checkValidity()) {
        render();
        return;
      }
      const value = input.valueAsNumber;
      if (!Number.isFinite(value)) return;
      if (key === 'rotation') {
        void options.update({ changes: { rotation: value } });
        return;
      }
      const positions = options
        .selection()
        .filter((shape) => shape.bounds)
        .map((shape) => {
          const bounds = { ...shape.bounds! };
          const next = Math.round(value * 360000);
          if ((key === 'w' || key === 'h') && shape.aspectRatioLocked && bounds[key] > 0) {
            const other = key === 'w' ? 'h' : 'w';
            bounds[other] = Math.max(1, Math.round((bounds[other] * next) / bounds[key]));
          }
          bounds[key] =
            key === 'w' || key === 'h' ? Math.max(shape.kind === 'connector' ? 0 : 1, next) : next;
          return { id: shape.id, bounds };
        });
      void options.update({ positions });
    };
  }
  const lockLabel = document.createElement('label');
  lockLabel.className = 'properties-lock';
  const lock = document.createElement('input');
  lock.type = 'checkbox';
  lockLabel.append(lock, 'Lock aspect ratio');
  pane.querySelector('[data-section="size"]')!.append(lockLabel);
  lock.onchange = () => void options.update({ changes: { aspectRatioLocked: lock.checked } });
  const textFields = pane.querySelector('[data-section="text"]')!;
  const alignmentLabel = document.createElement('label');
  alignmentLabel.textContent = 'Vertical alignment';
  const alignment = document.createElement('select');
  alignment.setAttribute('aria-label', 'Vertical alignment');
  for (const [value, title] of [
    ['top', 'Top'],
    ['center', 'Middle'],
    ['bottom', 'Bottom'],
    ['top-centered', 'Top Centered'],
    ['center-centered', 'Middle Centered'],
    ['bottom-centered', 'Bottom Centered'],
  ]) {
    alignment.add(new Option(title, value));
  }
  alignmentLabel.append(alignment);
  textFields.append(alignmentLabel);
  alignment.onchange = () =>
    void options.update({
      changes: {
        anchor: alignment.value.split('-')[0] as 'top' | 'center' | 'bottom',
        anchorCenter: alignment.value.endsWith('-centered'),
      },
    });
  const directionLabel = document.createElement('label');
  directionLabel.textContent = 'Text direction';
  const direction = document.createElement('select');
  direction.setAttribute('aria-label', 'Text direction');
  for (const [value, label] of [
    ['horz', 'Horizontal'],
    ['vert', 'Rotate all text 90°'],
    ['vert270', 'Rotate all text 270°'],
    ['wordArtVert', 'Stacked'],
  ])
    direction.add(new Option(label, value));
  directionLabel.append(direction);
  textFields.append(directionLabel);
  direction.onchange = () =>
    void options.update({
      changes: {
        direction: direction.value as NonNullable<NonNullable<EditCommand['changes']>['direction']>,
      },
    });
  const autoFitInputs = new Map<EditorShape['autoFit'], HTMLInputElement>();
  const autoFitGroup = document.createElement('div');
  autoFitGroup.setAttribute('role', 'radiogroup');
  autoFitGroup.setAttribute('aria-label', 'Autofit');
  for (const [mode, title] of [
    ['none', 'Do not Autofit'],
    ['normal', 'Shrink text on overflow'],
    ['shape', 'Resize shape to fit text'],
  ] as const) {
    const label = document.createElement('label');
    label.className = 'properties-lock';
    const input = document.createElement('input');
    input.type = 'radio';
    input.name = 'shape-text-autofit';
    input.value = mode;
    label.append(input, title);
    autoFitGroup.append(label);
    autoFitInputs.set(mode, input);
    input.onchange = () => {
      if (input.checked) void options.update({ changes: { autoFit: mode } });
    };
  }
  textFields.append(autoFitGroup);
  const margins = new Map<'left' | 'right' | 'top' | 'bottom', HTMLInputElement>();
  for (const key of ['left', 'right', 'top', 'bottom'] as const) {
    const label = document.createElement('label');
    const title = key[0]!.toUpperCase() + key.slice(1) + ' margin';
    label.textContent = title;
    const input = document.createElement('input');
    input.type = 'number';
    input.min = '0';
    input.max = '55.88';
    input.step = '0.01';
    input.setAttribute('aria-label', title);
    const unit = document.createElement('span');
    unit.className = 'properties-value';
    unit.append(input, 'cm');
    label.append(unit);
    textFields.append(label);
    margins.set(key, input);
    input.onkeydown = (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        input.blur();
      }
      if (event.key === 'Escape') {
        event.preventDefault();
        input.value = '';
        input.blur();
        render();
      }
    };
    input.onchange = () => {
      if (input.value === '' || !input.checkValidity()) {
        render();
        return;
      }
      void options.update({
        changes: { margins: { [key]: Math.round(input.valueAsNumber * 360000) } },
      });
    };
  }
  const wrapLabel = document.createElement('label');
  wrapLabel.className = 'properties-lock';
  const wrap = document.createElement('input');
  wrap.type = 'checkbox';
  wrapLabel.append(wrap, 'Wrap text in shape');
  textFields.append(wrapLabel);
  wrap.onchange = () => void options.update({ changes: { wrap: wrap.checked } });
  const columns = document.createElement('button');
  columns.textContent = 'Columns...';
  textFields.append(columns);
  columns.onclick = () => showColumns();
  function showColumns() {
    const shapes = options.selection();
    if (options.disabled() || !shapes.length || shapes.some((shape) => !shape.textFrame)) return;
    const dialog = document.createElement('dialog');
    dialog.className = 'section-dialog columns-dialog';
    dialog.setAttribute('aria-label', 'Columns');
    dialog.innerHTML = `<form><h2>Columns</h2>
      <label>Number of columns:<input name="count" type="number" min="1" max="16" step="1"></label>
      <label>Spacing between columns:<span class="properties-value"><input name="gap" type="number" min="0" max="40.64" step="0.01">cm</span></label>
      <div><button type="button">Cancel</button><button type="submit">OK</button></div></form>`;
    const count = dialog.querySelector<HTMLInputElement>('[name="count"]')!;
    const gap = dialog.querySelector<HTMLInputElement>('[name="gap"]')!;
    const counts = shapes.map((shape) => shape.textFrame!.columns?.count ?? 1);
    const gaps = shapes.map((shape) => shape.textFrame!.columns?.gapEmu ?? 0);
    count.value = counts.every((value) => value === counts[0]) ? String(counts[0]) : '';
    gap.value = gaps.every((value) => value === gaps[0])
      ? String(Number((gaps[0]! / 360000).toFixed(2)))
      : '';
    const initialCount = count.value;
    const initialGap = gap.value;
    dialog.querySelector('button')!.onclick = () => dialog.close();
    dialog.querySelector('form')!.onsubmit = (event) => {
      event.preventDefault();
      const textColumns: NonNullable<NonNullable<EditCommand['changes']>['textColumns']> = {};
      if (count.value !== initialCount && count.value !== '')
        textColumns.count = count.valueAsNumber;
      if (gap.value !== initialGap && gap.value !== '')
        textColumns.gapEmu = Math.round(gap.valueAsNumber * 360000);
      if (!Object.keys(textColumns).length) {
        dialog.close();
        return;
      }
      for (const control of dialog.querySelectorAll<HTMLButtonElement | HTMLInputElement>(
        'button,input',
      ))
        control.disabled = true;
      void options.update({ changes: { textColumns } }).then((ok) => {
        if (ok) dialog.close();
        else
          for (const control of dialog.querySelectorAll<HTMLButtonElement | HTMLInputElement>(
            'button,input',
          ))
            control.disabled = false;
      });
    };
    dialog.onclose = () => {
      dialog.remove();
      if (pane.hidden) options.focusCanvas();
      else columns.focus();
    };
    document.body.append(dialog);
    dialog.showModal();
    count.select();
  }
  pane.querySelector('button')!.onclick = () => {
    pane.hidden = true;
    document.querySelector('main')!.classList.remove('shape-properties-open');
    options.focusCanvas();
  };
  function render() {
    if (pane.hidden) return;
    const shapes = options.selection().filter((shape) => shape.bounds);
    const pending = options.disabled();
    paint.render(shapes, pending);
    const pictures = shapes.length > 0 && shapes.every((shape) => shape.kind === 'picture');
    const pictureTab = pane.querySelector<HTMLButtonElement>('#picture-tab')!;
    pictureTab.hidden = !pictures;
    if (!pictures && pictureTab.getAttribute('aria-selected') === 'true') selectTab(tabs[1]!);
    const title = pictures ? 'Format Picture' : 'Format Shape';
    pane.setAttribute('aria-label', title);
    pane.querySelector('h2')!.textContent = title;
    pane.querySelector('header button')!.setAttribute('aria-label', `Close ${title}`);
    for (const { input, key } of corrections) {
      input.disabled = pending || !pictures;
      if (pending || document.activeElement === input) continue;
      const values = shapes.map((shape) => Math.round((shape[key] ?? 0) * 100));
      input.value =
        pictures && values.every((value) => value === values[0]) ? String(values[0]) : '';
    }
    for (const input of transparency) {
      input.disabled = pending || !pictures;
      if (pending || document.activeElement === input) continue;
      const values = shapes.map((shape) => Math.round((1 - (shape.imageOpacity ?? 1)) * 100));
      input.value =
        pictures && values.every((value) => value === values[0]) ? String(values[0]) : '';
    }
    for (const [key, input] of fields) {
      input.disabled = pending || !shapes.length;
      if (pending) continue;
      const values = shapes.map((shape) =>
        key === 'rotation'
          ? shape.rotation
          : shape.bounds![key as keyof NonNullable<EditorShape['bounds']>] / 360000,
      );
      const common =
        values.length > 0 && values.every((value) => Math.abs(value - values[0]!) < 0.000001);
      if (document.activeElement !== input)
        input.value = common ? String(Number(values[0]!.toFixed(2))) : '';
    }
    lock.disabled = pending || !shapes.length;
    const frames = shapes.map((shape) => shape.textFrame);
    const textDisabled = pending || !frames.length || frames.some((frame) => !frame);
    for (const input of autoFitInputs.values()) input.disabled = textDisabled;
    direction.disabled = textDisabled;
    columns.disabled = textDisabled;
    alignment.disabled = textDisabled;
    wrap.disabled = textDisabled;
    for (const input of margins.values()) input.disabled = textDisabled;
    if (pending) return;
    for (const [mode, input] of autoFitInputs)
      input.checked = shapes.length > 0 && shapes.every((shape) => shape.autoFit === mode);
    direction.value =
      shapes.length && shapes.every((shape) => shape.textDirection === shapes[0]!.textDirection)
        ? shapes[0]!.textDirection
        : '';
    const validFrames = frames.filter(
      (frame): frame is NonNullable<typeof frame> => frame !== null,
    );
    alignment.value =
      validFrames.length &&
      validFrames.every((frame) => frame.anchor === validFrames[0]!.anchor) &&
      shapes.every((shape) => shape.anchorCenter === shapes[0]!.anchorCenter)
        ? validFrames[0]!.anchor + (shapes[0]!.anchorCenter ? '-centered' : '')
        : '';
    wrap.checked = validFrames.length > 0 && validFrames.every((frame) => frame.wrap);
    wrap.indeterminate = !wrap.checked && validFrames.some((frame) => frame.wrap);
    for (const [key, input] of margins) {
      if (document.activeElement === input) continue;
      input.value =
        validFrames.length &&
        validFrames.every((frame) => frame.margins[key] === validFrames[0]!.margins[key])
          ? String(Number((validFrames[0]!.margins[key] / 360000).toFixed(2)))
          : '';
    }
    lock.checked = shapes.length > 0 && shapes.every((shape) => shape.aspectRatioLocked);
    lock.indeterminate = !lock.checked && shapes.some((shape) => shape.aspectRatioLocked);
  }
  return {
    render,
    showColumns,
    showPicture() {
      pane.hidden = false;
      document.querySelector('main')!.classList.add('shape-properties-open');
      render();
      selectTab(pane.querySelector<HTMLButtonElement>('#picture-tab')!);
    },
    show() {
      pane.hidden = false;
      document.querySelector('main')!.classList.add('shape-properties-open');
      render();
    },
  };
}
