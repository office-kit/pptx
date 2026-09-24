import type { ArrowOptions } from '@office-kit/pptx';

const types = [
  ['none', 'No Arrow'],
  ['triangle', 'Triangle Arrow'],
  ['stealth', 'Stealth Arrow'],
  ['diamond', 'Diamond Arrow'],
  ['oval', 'Oval Arrow'],
  ['arrow', 'Open Arrow'],
] as const;
const sizes = ['sm', 'med', 'lg'] as const;

/** Small vector previews avoid font-dependent arrow glyphs. */
function preview(end: 'head' | 'tail', arrow: ArrowOptions): string {
  const w = { sm: 4, med: 7, lg: 10 }[arrow.width ?? 'med'];
  const len = { sm: 8, med: 13, lg: 18 }[arrow.length ?? 'med'];
  const tip = 51,
    base = tip - len;
  const marker = {
    none: '',
    triangle: `<path d="M${tip} 16 L${base} ${16 - w} V${16 + w} Z"/>`,
    stealth: `<path d="M${tip} 16 L${base} ${16 - w} L${base + len / 3} 16 L${base} ${16 + w} Z"/>`,
    diamond: `<path d="M${tip} 16 L${tip - len / 2} ${16 - w} L${base} 16 L${tip - len / 2} ${16 + w} Z"/>`,
    oval: `<ellipse cx="${tip - len / 2}" cy="16" rx="${len / 2}" ry="${w}"/>`,
    arrow: `<path d="M${base} ${16 - w} L${tip} 16 L${base} ${16 + w}" fill="none" stroke="currentColor" stroke-width="2"/>`,
  }[arrow.type];
  return `<svg viewBox="0 0 60 32" aria-hidden="true" focusable="false"><g fill="currentColor"${end === 'head' ? ' transform="translate(60 0) scale(-1 1)"' : ''}><path d="M8 16 H${arrow.type === 'none' || arrow.type === 'arrow' ? tip : base + len / 3}" fill="none" stroke="currentColor" stroke-width="2"/>${marker}</g></svg>`;
}

export function createArrowGallery(
  parent: Element,
  end: 'head' | 'tail',
  kind: 'type' | 'size',
  change: (options: Partial<ArrowOptions>) => void,
) {
  const title = `${end === 'head' ? 'Begin' : 'End'} Arrow ${kind}`;
  const row = document.createElement('div');
  row.className = 'properties-gallery-row';
  const label = document.createElement('span');
  label.textContent = title;
  const trigger = document.createElement('button');
  trigger.type = 'button';
  trigger.className = 'arrow-gallery-trigger';
  trigger.setAttribute('aria-label', title);
  trigger.setAttribute('aria-haspopup', 'listbox');
  trigger.setAttribute('aria-expanded', 'false');
  const popup = document.createElement('div');
  popup.className = 'arrow-gallery';
  popup.popover = 'auto';
  popup.setAttribute('role', 'listbox');
  popup.setAttribute('aria-label', title);
  popup.id = `arrow-${end}-${kind}-gallery`;
  trigger.setAttribute('aria-controls', popup.id);
  row.append(label, trigger, popup);
  parent.append(row);
  const entries: Array<{ value: string; label: string; options: Partial<ArrowOptions> }> =
    kind === 'type'
      ? types.map(([type, label]) => ({ value: type, label, options: { type } }))
      : sizes.flatMap((width, wi) =>
          sizes.map((length, li) => ({
            value: `${width}-${length}`,
            label: `${['Narrow', 'Medium', 'Wide'][wi]}, ${['Short', 'Medium', 'Long'][li]}`,
            options: { width, length },
          })),
        );
  const close = (restore: boolean) => {
    popup.hidePopover();
    if (restore) trigger.focus();
  };
  const buttons = entries.map((entry) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.setAttribute('role', 'option');
    button.setAttribute('aria-label', entry.label);
    button.title = entry.label;
    button.tabIndex = -1;
    button.onclick = () => {
      close(true);
      change(entry.options);
    };
    popup.append(button);
    return button;
  });
  const open = (last = false) => {
    if (trigger.disabled) return;
    popup.showPopover();
    const rect = trigger.getBoundingClientRect();
    popup.style.left = `${Math.max(4, Math.min(rect.right - popup.offsetWidth, innerWidth - popup.offsetWidth - 4))}px`;
    popup.style.top = `${Math.max(4, rect.bottom + popup.offsetHeight < innerHeight ? rect.bottom + 3 : rect.top - popup.offsetHeight - 3)}px`;
    (buttons.find((button) => button.getAttribute('aria-selected') === 'true') ??
      buttons[last ? buttons.length - 1 : 0])!.focus();
  };
  trigger.onclick = () => (popup.matches(':popover-open') ? close(false) : open());
  trigger.onkeydown = (event) => {
    if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      event.preventDefault();
      event.stopPropagation();
      open(event.key === 'ArrowUp');
    }
  };
  popup.addEventListener('toggle', () =>
    trigger.setAttribute('aria-expanded', String(popup.matches(':popover-open'))),
  );
  popup.onkeydown = (event) => {
    event.stopPropagation();
    if (event.key === 'Escape' || event.key === 'Tab') {
      if (event.key === 'Escape') event.preventDefault();
      close(true);
      return;
    }
    const current = buttons.indexOf(document.activeElement as HTMLButtonElement);
    const steps: Record<string, number> = {
      ArrowRight: 1,
      ArrowLeft: -1,
      ArrowDown: 3,
      ArrowUp: -3,
    };
    let next = current;
    if (event.key in steps)
      next = Math.max(0, Math.min(buttons.length - 1, current + steps[event.key]!));
    else if (event.key === 'Home') next = 0;
    else if (event.key === 'End') next = buttons.length - 1;
    else return;
    event.preventDefault();
    buttons[next]!.focus();
  };
  return {
    render(arrow: ArrowOptions, value: string, disabled: boolean) {
      trigger.disabled = disabled;
      if (disabled && popup.matches(':popover-open')) close(false);
      trigger.innerHTML = `${value ? preview(end, arrow) : '<span>—</span>'}<span aria-hidden="true">▾</span>`;
      trigger.title = entries.find((entry) => entry.value === value)?.label ?? 'Mixed';
      for (const [i, entry] of entries.entries()) {
        buttons[i]!.innerHTML = preview(end, { ...arrow, ...entry.options });
        buttons[i]!.setAttribute('aria-selected', String(entry.value === value));
      }
    },
  };
}
