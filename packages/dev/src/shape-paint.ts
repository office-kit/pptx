import { createArrowGallery } from './arrow-gallery.ts';
import type { EditorShape, EditCommand } from './editor.ts';

type Changes = NonNullable<EditCommand['changes']>;

/** Fill and line controls share the editor's selection and undo transaction. */
export function createShapePaint(
  parent: HTMLElement,
  update: (changes: Changes) => Promise<unknown>,
) {
  parent.innerHTML = `<details open><summary>Fill</summary><div class="properties-fields" data-paint="fill"></div></details>
    <details open><summary>Line</summary><div class="properties-fields" data-paint="line"></div></details>`;
  const modes = new Map<string, HTMLInputElement>();
  const colors = new Map<string, HTMLInputElement>();
  const transparency = new Map<string, [HTMLInputElement, HTMLInputElement]>();
  for (const section of ['fill', 'line'] as const) {
    const host = parent.querySelector<HTMLElement>(`[data-paint="${section}"]`)!;
    const group = document.createElement('div');
    group.setAttribute('role', 'radiogroup');
    group.setAttribute('aria-label', section === 'fill' ? 'Fill type' : 'Line type');
    for (const kind of ['none', 'solid'] as const) {
      const label = document.createElement('label');
      const input = document.createElement('input');
      input.type = 'radio';
      input.name = `shape-${section}-type`;
      input.value = kind;
      label.append(input, `${kind === 'none' ? 'No' : 'Solid'} ${section}`);
      group.append(label);
      modes.set(`${section}-${kind}`, input);
      input.onchange = () => {
        if (!input.checked) return;
        const color = colors.get(section)!.value;
        void update(
          section === 'fill'
            ? { fill: kind === 'none' ? 'none' : color }
            : { stroke: kind === 'none' ? 'none' : { color } },
        );
      };
    }
    host.append(group);
    const label = document.createElement('label');
    label.textContent = 'Color';
    const color = document.createElement('input');
    color.type = 'color';
    color.setAttribute('aria-label', section === 'fill' ? 'Fill color' : 'Line color');
    label.append(color);
    host.append(label);
    colors.set(section, color);
    color.onchange = () =>
      void update(section === 'fill' ? { fill: color.value } : { stroke: { color: color.value } });
  }
  for (const section of ['fill', 'line'] as const) {
    const label = document.createElement('label');
    label.textContent = 'Transparency';
    const slider = document.createElement('input');
    const number = document.createElement('input');
    slider.type = 'range';
    number.type = 'number';
    for (const input of [slider, number]) {
      input.min = '0';
      input.max = '100';
      input.step = '1';
      input.setAttribute(
        'aria-label',
        `${section === 'fill' ? 'Fill' : 'Line'} transparency${input === slider ? ' slider' : ''}`,
      );
      input.onchange = () => {
        if (input.value === '' || !input.checkValidity()) return;
        const opacity = 1 - input.valueAsNumber / 100;
        void update(section === 'fill' ? { fillOpacity: opacity } : { strokeOpacity: opacity });
      };
    }
    slider.style.width = '75px';
    slider.oninput = () => {
      number.value = slider.value;
    };
    const unit = document.createElement('span');
    unit.className = 'properties-value';
    unit.append(slider, number, '%');
    label.append(unit);
    parent.querySelector(`[data-paint="${section}"]`)!.append(label);
    transparency.set(section, [slider, number]);
  }
  const line = parent.querySelector('[data-paint="line"]')!;
  const widthLabel = document.createElement('label');
  widthLabel.textContent = 'Width';
  const width = document.createElement('input');
  width.type = 'number';
  width.min = '0';
  width.max = '1584';
  width.step = '0.25';
  width.setAttribute('aria-label', 'Line width');
  const unit = document.createElement('span');
  unit.className = 'properties-value';
  unit.append(width, 'pt');
  widthLabel.append(unit);
  line.append(widthLabel);
  width.onchange = () => {
    if (width.value && width.checkValidity())
      void update({ stroke: { widthEmu: Math.round(width.valueAsNumber * 12700) } });
  };
  const dashLabel = document.createElement('label');
  dashLabel.textContent = 'Dash type';
  const dash = document.createElement('select');
  dash.setAttribute('aria-label', 'Dash type');
  for (const [value, label] of [
    ['', ''],
    ['solid', 'Solid'],
    ['sysDot', 'Round Dot'],
    ['dot', 'Square Dot'],
    ['dash', 'Dash'],
    ['dashDot', 'Dash Dot'],
    ['lgDash', 'Long Dash'],
    ['lgDashDot', 'Long Dash Dot'],
    ['lgDashDotDot', 'Long Dash Dot Dot'],
    ['sysDash', 'System Dash'],
    ['sysDashDot', 'System Dash Dot'],
    ['sysDashDotDot', 'System Dash Dot Dot'],
  ])
    dash.add(new Option(label, value));
  dashLabel.append(dash);
  line.append(dashLabel);
  dash.onchange = () => {
    if (dash.value) void update({ strokeDash: dash.value as NonNullable<Changes['strokeDash']> });
  };
  const lineStyles = [
    {
      key: 'compound',
      command: 'strokeCompound',
      label: 'Compound type',
      values: [
        ['sng', 'Single'],
        ['dbl', 'Double'],
        ['thickThin', 'Thick Thin'],
        ['thinThick', 'Thin Thick'],
        ['tri', 'Triple'],
      ],
    },
    {
      key: 'cap',
      command: 'strokeCap',
      label: 'Cap type',
      values: [
        ['flat', 'Flat'],
        ['rnd', 'Round'],
        ['sq', 'Square'],
      ],
    },
    {
      key: 'join',
      command: 'strokeJoin',
      label: 'Join type',
      values: [
        ['round', 'Round'],
        ['bevel', 'Bevel'],
        ['miter', 'Miter'],
      ],
    },
  ] as const;
  const styleInputs = new Map<string, HTMLSelectElement>();
  for (const setting of lineStyles) {
    const label = document.createElement('label');
    label.textContent = setting.label;
    const input = document.createElement('select');
    input.setAttribute('aria-label', setting.label);
    input.add(new Option('', ''));
    for (const [value, title] of setting.values) input.add(new Option(title, value));
    input.onchange = () => {
      if (input.value) void update({ [setting.command]: input.value });
    };
    label.append(input);
    if (setting.key === 'compound') line.insertBefore(label, dashLabel);
    else line.append(label);
    styleInputs.set(setting.key, input);
  }
  const arrowInputs = new Map<
    string,
    [ReturnType<typeof createArrowGallery>, ReturnType<typeof createArrowGallery>]
  >();
  for (const [end, command] of [
    ['head', 'strokeHeadArrow'],
    ['tail', 'strokeTailArrow'],
  ] as const) {
    const change = (options: NonNullable<Changes[typeof command]>) => {
      void update({ [command]: options });
    };
    arrowInputs.set(`${end}Arrow`, [
      createArrowGallery(line, end, 'type', change),
      createArrowGallery(line, end, 'size', change),
    ]);
  }
  return {
    render(shapes: EditorShape[], pending: boolean) {
      const disabled = pending || !shapes.length || shapes.some((shape) => !shape.paint);
      for (const input of parent.querySelectorAll<HTMLInputElement | HTMLSelectElement>(
        'input,select',
      ))
        input.disabled = disabled;
      for (const end of ['headArrow', 'tailArrow'] as const) {
        const [type, size] = arrowInputs.get(end)!;
        const arrows = shapes.map((shape) => shape.paint?.[end]);
        const types = arrows.map((arrow) => arrow?.type ?? 'none');
        const sizes = arrows.map((arrow) => `${arrow?.width ?? 'med'}-${arrow?.length ?? 'med'}`);
        const commonType =
          types.length && types.every((value) => value === types[0]) ? types[0]! : '';
        const commonSize =
          sizes.length && sizes.every((value) => value === sizes[0]) ? sizes[0]! : '';
        const arrow = arrows[0] ?? { type: 'none' as const };
        const noLine = shapes.every((shape) => shape.paint?.stroke.kind === 'none');
        type.render(arrow, commonType, disabled || noLine);
        size.render(
          arrow,
          commonSize,
          disabled || noLine || types.some((value) => value === 'none'),
        );
      }
      if (pending) return;
      for (const section of ['fill', 'line'] as const) {
        const paints = shapes.map((shape) =>
          section === 'fill' ? shape.paint?.fill : shape.paint?.stroke,
        );
        for (const kind of ['none', 'solid'])
          modes.get(`${section}-${kind}`)!.checked =
            paints.length > 0 && paints.every((paint) => paint?.kind === kind);
        const opacityValues = shapes.map((shape) =>
          section === 'fill' ? shape.paint?.fillOpacity : shape.paint?.strokeOpacity,
        );
        const [slider, number] = transparency.get(section)!;
        const opacityDisabled = disabled || opacityValues.some((value) => value == null);
        slider.disabled = number.disabled = opacityDisabled;
        const commonOpacity =
          opacityValues.length &&
          opacityValues[0] != null &&
          opacityValues.every((value) => value === opacityValues[0]);
        if (document.activeElement !== number)
          number.value = commonOpacity ? String(Math.round((1 - opacityValues[0]!) * 100)) : '';
        if (document.activeElement !== slider) slider.value = number.value || '0';
        const color = colors.get(section)!;
        const values = paints.map((paint) => (paint?.kind === 'solid' ? paint.color : null));
        const common = values.length && values.every((value) => value === values[0]);
        if (document.activeElement !== color)
          color.value = common && /^#[0-9a-f]{6}$/i.test(values[0] ?? '') ? values[0]! : '#000000';
        color.title = common && values[0] ? values[0] : 'Mixed or inherited color';
        color.disabled = disabled || paints.every((paint) => paint?.kind === 'none');
        if (
          section === 'fill' &&
          shapes.some((shape) => shape.kind === 'connector' || shape.kind === 'picture')
        ) {
          color.disabled = true;
          for (const kind of ['none', 'solid']) modes.get(`${section}-${kind}`)!.disabled = true;
        }
      }
      const strokes = shapes.map((shape) => shape.paint?.stroke);
      const widths = strokes.map((stroke) =>
        stroke?.kind === 'solid' ? (stroke.widthEmu ?? 9525) / 12700 : null,
      );
      if (document.activeElement !== width)
        width.value =
          widths.length && widths[0] !== null && widths.every((value) => value === widths[0])
            ? String(widths[0])
            : '';
      const noLine = strokes.every((stroke) => stroke?.kind === 'none');
      for (const setting of lineStyles) {
        const input = styleInputs.get(setting.key)!;
        input.disabled = disabled || noLine;
        const values = shapes.map((shape) => shape.paint?.[setting.key]);
        input.value =
          values.length && values.every((value) => value === values[0]) ? (values[0] ?? '') : '';
      }
      width.disabled = disabled || noLine;
      dash.disabled = disabled || noLine;
      dash.value =
        shapes.length &&
        shapes.every(
          (shape) => (shape.paint?.dash ?? 'solid') === (shapes[0]!.paint?.dash ?? 'solid'),
        )
          ? (shapes[0]!.paint?.dash ?? 'solid')
          : '';
    },
  };
}
