<script lang="ts">
  // Hand-tuned quick controls for the most-used shape properties, shown at the
  // top of the properties panel. These are ergonomic shortcuts; the full,
  // exhaustive list still lives below in the auto-generated section, so nothing
  // here is the *only* path to a capability.
  import { getEditor } from '../core/context.ts';
  import {
    getShapeBounds,
    getShapeRotation,
    getShapeText,
    inches,
    setShapeBounds,
    setShapeFill,
    setShapeRotation,
    setShapeStroke,
    setShapeText,
  } from '@office-kit/pptx';
  import { selectedShapeId } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;

  const shape = $derived.by(() => {
    doc.version;
    const sel = doc.selection;
    if (sel.kind !== 'shape') return null;
    const id = selectedShapeId(sel);
    return id == null ? null : doc.shapeById(sel.slideIndex, id);
  });

  function emuToIn(v: number | undefined): number {
    const per = inches(1) as unknown as number;
    return v == null ? 0 : Math.round((v / per) * 100) / 100;
  }
  function inToEmu(v: number): number {
    return Math.round((inches(1) as unknown as number) * v);
  }

  const bounds = $derived.by(() => {
    const s = shape;
    if (!s) return null;
    try {
      const b = getShapeBounds(s);
      if (!b) return null;
      return {
        x: emuToIn(b.x as unknown as number),
        y: emuToIn(b.y as unknown as number),
        w: emuToIn(b.w as unknown as number),
        h: emuToIn(b.h as unknown as number),
      };
    } catch {
      return null;
    }
  });

  const rotation = $derived.by(() => {
    const s = shape;
    if (!s) return 0;
    try {
      return getShapeRotation(s);
    } catch {
      return 0;
    }
  });

  const text = $derived.by(() => {
    const s = shape;
    if (!s) return '';
    try {
      return getShapeText(s);
    } catch {
      return '';
    }
  });

  let fill = $state('#3b6ea5');
  let strokeColor = $state('#1f1f1f');

  function applyFill() {
    editor.invoke('setShapeFill', { color: fill.replace('#', '') });
  }
  function applyStroke() {
    // setShapeStroke(shape, options: { color?, widthEmu? }) — pass the object.
    editor.invoke('setShapeStroke', { options: { color: strokeColor.replace('#', '') } });
  }
  function setBoundsField(field: 'x' | 'y' | 'w' | 'h', value: number) {
    const s = shape;
    const b = bounds;
    if (!s || !b) return;
    const next = { ...b, [field]: value };
    doc.transact('Set bounds', () =>
      setShapeBounds(s, {
        x: inToEmu(next.x) as never,
        y: inToEmu(next.y) as never,
        w: inToEmu(next.w) as never,
        h: inToEmu(next.h) as never,
      }),
    );
  }
  function applyRotation(deg: number) {
    const s = shape;
    if (!s) return;
    doc.transact('Rotate', () => setShapeRotation(s, deg));
  }
  function applyText(value: string) {
    const s = shape;
    if (!s) return;
    doc.transact('Edit text', () => setShapeText(s, value));
  }
</script>

{#if shape}
  <div class="bespoke">
    <div class="sec">
      <div class="sec-title">{t('Fill & outline')}</div>
      <div class="row2">
        <label class="mini">
          <span>{t('Fill')}</span>
          <span class="colorwrap">
            <input type="color" bind:value={fill} onchange={applyFill} />
          </span>
        </label>
        <label class="mini">
          <span>{t('Outline')}</span>
          <span class="colorwrap">
            <input type="color" bind:value={strokeColor} onchange={applyStroke} />
          </span>
        </label>
      </div>
    </div>

    {#if bounds}
      <div class="sec">
        <div class="sec-title">{t('Position & size (in)')}</div>
        <div class="grid4">
          <label class="mini"><span>X</span>
            <input class="ok-input" type="number" step="0.01" value={bounds.x}
              onchange={(e) => setBoundsField('x', Number(e.currentTarget.value))} /></label>
          <label class="mini"><span>Y</span>
            <input class="ok-input" type="number" step="0.01" value={bounds.y}
              onchange={(e) => setBoundsField('y', Number(e.currentTarget.value))} /></label>
          <label class="mini"><span>W</span>
            <input class="ok-input" type="number" step="0.01" value={bounds.w}
              onchange={(e) => setBoundsField('w', Number(e.currentTarget.value))} /></label>
          <label class="mini"><span>H</span>
            <input class="ok-input" type="number" step="0.01" value={bounds.h}
              onchange={(e) => setBoundsField('h', Number(e.currentTarget.value))} /></label>
        </div>
      </div>
    {/if}

    <div class="sec">
      <div class="sec-title">{t('Rotation')}</div>
      <div class="rotrow">
        <input class="ok-input" type="number" step="1" value={rotation}
          onchange={(e) => applyRotation(Number(e.currentTarget.value))} />
        <span class="deg">°</span>
      </div>
    </div>

    <div class="sec">
      <div class="sec-title">{t('Text')}</div>
      <textarea class="ok-input" rows="2" value={text}
        onchange={(e) => applyText(e.currentTarget.value)}></textarea>
    </div>
  </div>
{/if}

<style>
  .bespoke {
    padding: 8px 10px;
    border-bottom: 1px solid var(--ok-border);
    display: flex;
    flex-direction: column;
    gap: 10px;
  }
  .sec-title {
    font-size: 11px;
    font-weight: 600;
    color: var(--ok-text-2);
    margin-bottom: 5px;
  }
  .row2 {
    display: flex;
    gap: 10px;
  }
  .mini {
    display: flex;
    flex-direction: column;
    gap: 3px;
    font-size: 11px;
    color: var(--ok-text-2);
    flex: 1;
  }
  .colorwrap input[type='color'] {
    width: 100%;
    height: 26px;
    border: 1px solid var(--ok-border-strong);
    border-radius: var(--ok-radius);
    background: none;
    padding: 0;
    cursor: pointer;
  }
  .grid4 {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 6px;
  }
  .rotrow {
    display: flex;
    align-items: center;
    gap: 6px;
  }
  .rotrow .ok-input {
    width: 80px;
  }
  .deg {
    color: var(--ok-text-2);
  }
  textarea.ok-input {
    resize: vertical;
    width: 100%;
    font-family: var(--ok-font);
  }
</style>
