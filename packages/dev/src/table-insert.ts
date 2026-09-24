type TableSize = { rows: number; columns: number };
export function openTablePicker(
  origin: HTMLElement,
  insert: (size: TableSize) => Promise<boolean>,
) {
  document.querySelector('.table-picker')?.remove();
  const picker = document.createElement('div');
  picker.className = 'table-picker';
  picker.popover = 'auto';
  picker.innerHTML =
    '<div role="status" aria-live="polite">1 × 1 Table</div><div role="grid" aria-label="Table size"></div><button type="button" data-dialog>Insert Table…</button>';
  const grid = picker.querySelector<HTMLElement>('[role=grid]')!;
  const status = picker.querySelector<HTMLElement>('[role=status]')!;
  const cells: HTMLButtonElement[] = [];
  let rows = 1,
    columns = 1;
  const update = (r: number, c: number, focus = false) => {
    rows = r;
    columns = c;
    status.textContent = `${columns} × ${rows} Table`;
    cells.forEach((cell, index) => {
      const row = Math.floor(index / 10) + 1,
        column = (index % 10) + 1;
      cell.setAttribute('aria-selected', String(row <= rows && column <= columns));
      cell.tabIndex = row === rows && column === columns ? 0 : -1;
      if (focus && cell.tabIndex === 0) cell.focus();
    });
  };
  const close = (restore = false) => {
    picker.hidePopover();
    picker.remove();
    origin.setAttribute('aria-expanded', 'false');
    if (restore) origin.focus();
  };
  const choose = () => {
    close();
    void insert({ rows, columns });
  };
  for (let r = 1; r <= 8; r++) {
    const row = document.createElement('div');
    row.setAttribute('role', 'row');
    for (let c = 1; c <= 10; c++) {
      const cell = document.createElement('button');
      cell.type = 'button';
      cell.setAttribute('role', 'gridcell');
      cell.setAttribute('aria-label', `${c} columns, ${r} rows`);
      cell.onpointerenter = () => update(r, c);
      cell.onfocus = () => update(r, c);
      cell.onclick = () => {
        update(r, c);
        choose();
      };
      cells.push(cell);
      row.append(cell);
    }
    grid.append(row);
  }
  picker.onkeydown = (event) => {
    event.stopPropagation();
    if (event.key === 'Escape') {
      event.preventDefault();
      close(true);
      return;
    }
    if (!grid.contains(event.target as Node)) return;
    const next = {
      ArrowLeft: [rows, Math.max(1, columns - 1)],
      ArrowRight: [rows, Math.min(10, columns + 1)],
      ArrowUp: [Math.max(1, rows - 1), columns],
      ArrowDown: [Math.min(8, rows + 1), columns],
      Home: [rows, 1],
      End: [rows, 10],
    }[event.key];
    if (next) {
      event.preventDefault();
      update(next[0]!, next[1]!, true);
    }
  };
  picker.querySelector<HTMLButtonElement>('[data-dialog]')!.onclick = () => {
    close();
    openInsertTableDialog(origin, insert);
  };
  picker.addEventListener('toggle', (event) => {
    if ((event as ToggleEvent).newState === 'closed') {
      origin.setAttribute('aria-expanded', 'false');
      picker.remove();
    }
  });
  origin.setAttribute('aria-haspopup', 'grid');
  origin.setAttribute('aria-expanded', 'true');
  document.body.append(picker);
  picker.showPopover();
  const rect = origin.getBoundingClientRect();
  picker.style.left = `${Math.max(4, Math.min(rect.left, innerWidth - picker.offsetWidth - 4))}px`;
  picker.style.top = `${Math.max(4, Math.min(rect.bottom, innerHeight - picker.offsetHeight - 4))}px`;
  update(1, 1, true);
}
function openInsertTableDialog(origin: HTMLElement, insert: (size: TableSize) => Promise<boolean>) {
  const dialog = document.createElement('dialog');
  dialog.className = 'section-dialog insert-table-dialog';
  dialog.setAttribute('aria-label', 'Insert Table');
  dialog.innerHTML =
    '<form><h2>Insert Table</h2><label>Number of columns:<input name="columns" type="number" min="1" max="75" step="1" value="5" required autofocus></label><label>Number of rows:<input name="rows" type="number" min="1" max="75" step="1" value="2" required></label><p role="alert" hidden>Could not insert the table. Please try again.</p><div><button type="button" data-cancel>Cancel</button><button type="submit">OK</button></div></form>';
  let saving = false,
    saved = false;
  dialog.oncancel = (event) => {
    if (saving) event.preventDefault();
  };
  dialog.onclose = () => {
    dialog.remove();
    if (!saved) origin.focus();
  };
  dialog.querySelector<HTMLButtonElement>('[data-cancel]')!.onclick = () => dialog.close();
  dialog.querySelector('form')!.onsubmit = async (event) => {
    event.preventDefault();
    if (saving) return;
    const rows = dialog.querySelector<HTMLInputElement>('[name=rows]')!.valueAsNumber;
    const columns = dialog.querySelector<HTMLInputElement>('[name=columns]')!.valueAsNumber;
    saving = true;
    const controls = dialog.querySelectorAll<HTMLInputElement | HTMLButtonElement>('input,button');
    controls.forEach((control) => {
      control.disabled = true;
    });
    try {
      saved = await insert({ rows, columns });
      if (saved) dialog.close();
      else dialog.querySelector<HTMLElement>('[role=alert]')!.hidden = false;
    } finally {
      saving = false;
      controls.forEach((control) => {
        control.disabled = false;
      });
    }
  };
  document.body.append(dialog);
  dialog.showModal();
  dialog.querySelector<HTMLInputElement>('[name=columns]')!.select();
}
