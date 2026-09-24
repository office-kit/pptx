import type { EditCommand } from './editor.ts';
type Size = NonNullable<EditCommand['size']>;
// Unitless input uses the unit displayed beside the field (centimeters).
function dimension(value: string): number | null {
  const match = /^\s*(\d+(?:\.\d*)?|\.\d+)\s*(cm|in)?\s*$/i.exec(value);
  if (!match) return null;
  const emu = Number(match[1]) * (match[2]?.toLowerCase() === 'in' ? 914400 : 360000);
  if (!Number.isFinite(emu) || emu < 914400 || emu > 51206400) return null;
  return Math.round(emu);
}
function centimeters(emu: number): string {
  return String(Number((emu / 360000).toFixed(6)));
}
export const slideSizePresets: Array<{ label: string; size: Size }> = [
  { label: 'Standard (4:3)', size: { width: 9144000, height: 6858000, type: 'screen4x3' } },
  { label: 'Widescreen (16:9)', size: { width: 12192000, height: 6858000, type: 'screen16x9' } },
];
// Microsoft’s Mac Page Setup table lists the slide canvas dimensions, which
// differ from the physical paper dimensions shown in several labels.
const preset = (label: string, width: number, height: number, type: string) => ({
  label,
  size: { width: Math.round(width * 914400), height: Math.round(height * 914400), type },
});
const presets = [
  preset('On-screen Show (4:3)', 10, 7.5, 'screen4x3'),
  preset('Letter Paper (8.5x11 in)', 10, 7.5, 'letter'),
  preset('Ledger Paper (11x17 in)', 13.319, 9.99, 'ledger'),
  preset('A3 Paper (297x420 mm)', 14, 10.5, 'A3'),
  preset('A4 Paper (210x297 mm)', 10.833, 7.5, 'A4'),
  preset('B4 (ISO) Paper (250x353 mm)', 11.84, 8.88, 'B4ISO'),
  preset('B5 (ISO) Paper (176x250 mm)', 7.84, 5.88, 'B5ISO'),
  preset('35mm Slides', 11.25, 7.5, '35mm'),
  preset('Overhead', 10, 7.5, 'overhead'),
  preset('Banner', 8, 1, 'banner'),
  preset('On-screen Show (16:9)', 10, 5.625, 'screen16x9'),
  preset('On-screen Show (16:10)', 10, 6.25, 'screen16x10'),
  preset('Widescreen', 40 / 3, 7.5, 'screen16x9'),
];
function dialog(label: string, content: string) {
  const element = document.createElement('dialog');
  element.className = 'section-dialog page-setup-dialog';
  element.setAttribute('aria-label', label);
  element.innerHTML = `<form><h2>${label}</h2>${content}</form>`;
  const restoreFocus = document.activeElement;
  element.onclose = () => {
    element.remove();
    if (restoreFocus instanceof HTMLElement && restoreFocus.isConnected) restoreFocus.focus();
  };
  element.querySelector<HTMLButtonElement>('[data-cancel]')!.onclick = () => element.close();
  document.body.append(element);
  element.showModal();
  return element;
}
export function confirmSlideSize(
  size: Size,
  save: (size: Size, scale: boolean) => Promise<boolean>,
) {
  const element = dialog(
    'Scale Slide Content',
    `<p>Would you like to scale the content to fit the new slide size?</p><div><button type="button" data-cancel>Cancel</button><button type="submit" value="none">Don't Scale</button><button type="submit" value="scale" autofocus>Scale</button></div><p role="alert" hidden>Could not change the slide size. Please try again.</p>`,
  );
  let saving = false;
  element.oncancel = (event) => {
    if (saving) event.preventDefault();
  };
  element.querySelector('form')!.onsubmit = async (event) => {
    event.preventDefault();
    if (saving) return;
    saving = true;
    const scale = (event.submitter as HTMLButtonElement | null)?.value !== 'none';
    for (const button of element.querySelectorAll('button')) button.disabled = true;
    try {
      if (await save(size, scale)) element.close();
      else element.querySelector<HTMLElement>('[role=alert]')!.hidden = false;
    } finally {
      saving = false;
      for (const button of element.querySelectorAll('button')) button.disabled = false;
    }
  };
}
export function openPageSetup(
  current: Size,
  save: (size: Size, scale: boolean) => Promise<boolean>,
) {
  const element = dialog(
    'Page Setup',
    `<label>Slides sized for:<select name="preset">${presets.map((preset, i) => `<option value="${i}">${preset.label}</option>`).join('')}<option value="custom">Custom</option></select></label><label>Width:<span><input name="width" type="text" required> cm</span></label><label>Height:<span><input name="height" type="text" required> cm</span></label><label>Number slides from:<input name="firstSlideNumber" type="number" min="0" max="9999" step="1" required></label><fieldset><legend>Slides</legend><label><input type="radio" name="orientation" value="landscape"> Landscape</label><label><input type="radio" name="orientation" value="portrait"> Portrait</label></fieldset><fieldset><legend>Notes, handouts &amp; outline</legend><label><input type="radio" name="notesOrientation" value="landscape"> Landscape</label><label><input type="radio" name="notesOrientation" value="portrait"> Portrait</label></fieldset><div><button type="button" data-cancel>Cancel</button><button type="submit">OK</button></div><p role="alert" hidden>Could not change the slide size. Please try again.</p>`,
  );
  const form = element.querySelector('form')!;
  const width = form.elements.namedItem('width') as HTMLInputElement;
  const height = form.elements.namedItem('height') as HTMLInputElement;
  const preset = form.elements.namedItem('preset') as HTMLSelectElement;
  const firstSlideNumber = form.elements.namedItem('firstSlideNumber') as HTMLInputElement;
  firstSlideNumber.value = String(current.firstSlideNumber ?? 1);
  const notesPortrait = form.querySelector<HTMLInputElement>(
    '[name=notesOrientation][value=portrait]',
  )!;
  const notesLandscape = form.querySelector<HTMLInputElement>(
    '[name=notesOrientation][value=landscape]',
  )!;
  notesLandscape.checked = current.notesOrientation === 'landscape';
  notesPortrait.checked = !notesLandscape.checked;
  const portrait = form.querySelector<HTMLInputElement>('[name=orientation][value=portrait]')!;
  const landscape = form.querySelector<HTMLInputElement>('[name=orientation][value=landscape]')!;
  const orientation = () => {
    const w = dimension(width.value),
      h = dimension(height.value);
    if (w !== null && h !== null) {
      portrait.checked = w < h;
      landscape.checked = !portrait.checked;
    }
  };
  const setSize = (size: Size) => {
    width.value = centimeters(size.width);
    height.value = centimeters(size.height);
    width.setCustomValidity('');
    height.setCustomValidity('');
    orientation();
  };
  const validate = (input: HTMLInputElement) => {
    const value = dimension(input.value);
    input.setCustomValidity(
      value === null ? 'Enter a size from 2.54 to 142.24 cm (1 to 56 in).' : '',
    );
    return value;
  };
  setSize(current);
  const initial = presets.findIndex(
    ({ size }) =>
      (current.type === undefined || size.type === current.type) &&
      ((size.width === current.width && size.height === current.height) ||
        (size.height === current.width && size.width === current.height)),
  );
  preset.value = initial < 0 ? 'custom' : String(initial);
  preset.onchange = () => {
    if (preset.value !== 'custom') {
      const size = presets[Number(preset.value)]!.size;
      setSize(portrait.checked ? { ...size, width: size.height, height: size.width } : size);
    }
  };
  for (const input of [width, height]) {
    input.oninput = () => {
      preset.value = 'custom';
      validate(input);
      orientation();
    };
    input.onblur = () => {
      const value = validate(input);
      if (value !== null) input.value = centimeters(value);
    };
  }
  for (const input of [portrait, landscape])
    input.onchange = () => {
      const w = validate(width),
        h = validate(height);
      if (w !== null && h !== null && w < h !== portrait.checked) {
        [width.value, height.value] = [height.value, width.value];
      }
    };
  let saving = false;
  element.oncancel = (event) => {
    if (saving) event.preventDefault();
  };
  form.onsubmit = async (event) => {
    event.preventDefault();
    if (saving) return;
    const w = validate(width),
      h = validate(height);
    if (w === null || h === null) {
      form.reportValidity();
      return;
    }
    const size: Size = {
      width: w,
      height: h,
      firstSlideNumber: Number(firstSlideNumber.value),
      notesOrientation: notesLandscape.checked ? 'landscape' : 'portrait',
      type:
        preset.value === 'custom'
          ? 'custom'
          : (presets[Number(preset.value)]!.size.type ?? 'custom'),
    };
    if (size.width !== current.width || size.height !== current.height) {
      element.close();
      confirmSlideSize(size, save);
    } else if (
      size.type !== current.type ||
      size.firstSlideNumber !== (current.firstSlideNumber ?? 1) ||
      size.notesOrientation !== (current.notesOrientation ?? 'portrait')
    ) {
      saving = true;
      for (const button of element.querySelectorAll('button')) button.disabled = true;
      try {
        if (await save(size, false)) element.close();
        else element.querySelector<HTMLElement>('[role=alert]')!.hidden = false;
      } finally {
        saving = false;
        for (const button of element.querySelectorAll('button')) button.disabled = false;
      }
    } else element.close();
  };
}
