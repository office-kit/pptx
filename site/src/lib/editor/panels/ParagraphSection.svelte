<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import { getShapeKind, hasShapeText, isShapeTextBox, getShapeParagraphCount, getShapeParagraphElements, getParagraphPropertiesEffective, setParagraphAlignment, setParagraphBullet, setParagraphLevel, setParagraphLineSpacing, setParagraphSpacing } from '@office-kit/pptx';

  const doc = getEditor().doc;
  let target = $state<{ key: string; index: number } | null>(null);
  const current = $derived.by(() => {
    doc.version;
    const sel = doc.selection;
    if (sel.kind !== 'shape' || sel.shapeIds.length !== 1) return null;
    const shape = doc.shapeById(sel.slideIndex, sel.shapeIds[0]!);
    if (!shape || getShapeKind(shape) !== 'shape' || (!hasShapeText(shape) && !isShapeTextBox(shape))) return null;
    const paragraphs = Array.from({ length: getShapeParagraphCount(shape) }, (_, index) => ({
      index,
      text: getShapeParagraphElements(shape, index).map(element => element.kind === 'br' ? ' ' : element.text).join(''),
      properties: getParagraphPropertiesEffective(doc.pres, shape, index),
    }));
    return { shape, key: `${sel.slideIndex}:${sel.shapeIds[0]}`, paragraphs };
  });
  const index = $derived(current && target?.key === current.key && target.index < current.paragraphs.length ? target.index : -1);
  const selected = $derived(current?.paragraphs.filter(p => index === -1 || p.index === index) ?? []);
  function common<T>(read: (p: NonNullable<typeof current>['paragraphs'][number]['properties']) => T): T | undefined {
    const values = selected.map(p => read(p.properties));
    return values.every(value => value === values[0]) ? values[0] : undefined;
  }
  const align = $derived(common(p => p.align ?? 'left'));
  const bullet = $derived(common(p => typeof p.bullet === 'string' ? p.bullet : p.bullet === null ? 'none' : 'custom'));
  const level = $derived(common(p => p.level));
  const lineKind = $derived(common(p => p.lineSpacing?.kind ?? 'inherit'));
  const lineValue = $derived(common(p => p.lineSpacing?.value));
  const before = $derived(common(p => p.spcBefPts));
  const after = $derived(common(p => p.spcAftPts));

  function apply(edit: (shape: NonNullable<typeof current>['shape'], index: number) => void) {
    if (!current) return;
    const shape = current.shape;
    doc.transact(t('Format paragraphs'), () => {
      for (const p of selected) edit(shape, p.index);
    });
  }
  function alignment(value: string) {
    if (value === 'left' || value === 'center' || value === 'right' || value === 'justify') apply((s, i) => setParagraphAlignment(s, i, value));
  }
  function bullets(value: string) {
    if (value === 'none' || value === 'bullet' || value === 'number') apply((s, i) => setParagraphBullet(s, i, value));
  }
  function spacingKind(value: string) {
    if (value === 'inherit') apply((s, i) => setParagraphLineSpacing(s, i, null));
    else if (value === 'pct' || value === 'pts') apply((s, i) => setParagraphLineSpacing(s, i, { kind: value, value: value === 'pct' ? 1 : 18 }));
  }
  function lineSpacing(input: HTMLInputElement) {
    if (!input.reportValidity() || !Number.isFinite(input.valueAsNumber)) return;
    const kind = lineKind;
    if (kind === 'pct' || kind === 'pts') apply((s, i) => setParagraphLineSpacing(s, i, { kind, value: input.valueAsNumber }));
  }
  function paragraphSpacing(input: HTMLInputElement, side: 'beforePts' | 'afterPts') {
    if (!input.reportValidity()) return;
    const value = input.value === '' ? null : input.valueAsNumber;
    if (value !== null && !Number.isFinite(value)) return;
    apply((s, i) => setParagraphSpacing(s, i, { [side]: value }));
  }
</script>

{#if current}
  <section class="paragraphs" aria-label={t('Paragraph formatting')}>
    <strong>{t('Paragraph formatting')}</strong>
    <label>{t('Apply to paragraphs')}
      <select aria-label={t('Apply to paragraphs')} value={index} onchange={e => target = { key: current!.key, index: Number(e.currentTarget.value) }}>
        <option value={-1}>{t('All paragraphs')}</option>
        {#each current.paragraphs as paragraph}
          <option value={paragraph.index}>{paragraph.index + 1}: {paragraph.text.slice(0, 50) || t('Empty paragraph')}</option>
        {/each}
      </select>
    </label>
    <div class="grid">
      <label>{t('Paragraph alignment')}
        <select aria-label={t('Paragraph alignment')} value={align ?? ''} onchange={e => alignment(e.currentTarget.value)}>
          <option value="" disabled>{t('Mixed')}</option>
          <option value="left">{t('Left')}</option><option value="center">{t('Center')}</option><option value="right">{t('Right')}</option><option value="justify">{t('Justify')}</option>
          {#if align && !['left', 'center', 'right', 'justify'].includes(align)}<option value={align}>{t('Custom')}</option>{/if}
        </select>
      </label>
      <label>{t('List style')}
        <select aria-label={t('List style')} value={bullet ?? ''} onchange={e => bullets(e.currentTarget.value)}>
          <option value="" disabled>{t('Mixed')}</option><option value="none">{t('No list')}</option><option value="bullet">{t('Bulleted list')}</option><option value="number">{t('Numbered list')}</option>
          {#if bullet === 'custom'}<option value="custom" disabled>{t('Custom')}</option>{/if}
        </select>
      </label>
      <label>{t('List level')}
        <select aria-label={t('List level')} value={level ?? ''} onchange={e => { const value = Number(e.currentTarget.value); apply((s, i) => setParagraphLevel(s, i, value)); }}>
          <option value="" disabled>{t('Mixed')}</option>
          {#each Array.from({ length: 9 }, (_, i) => i) as value}<option value={value}>{value + 1}</option>{/each}
        </select>
      </label>
      <label>{t('Line spacing mode')}
        <select aria-label={t('Line spacing mode')} value={lineKind ?? ''} onchange={e => spacingKind(e.currentTarget.value)}>
          <option value="" disabled>{t('Mixed')}</option><option value="inherit">{t('Inherit')}</option><option value="pct">{t('Multiple')}</option><option value="pts">{t('Points')}</option>
        </select>
      </label>
      {#if lineKind === 'pct' || lineKind === 'pts'}
        <label>{t('Line spacing value')}
          <input type="number" min="0" step="0.01" required value={lineValue ?? ''} placeholder={t('Mixed')} onchange={e => lineSpacing(e.currentTarget)} />
        </label>
      {/if}
      <label>{t('Before paragraph (pt)')}
        <input type="number" min="0" step="0.01" value={before ?? ''} placeholder={t(before === undefined ? 'Mixed' : 'Inherit')} onchange={e => paragraphSpacing(e.currentTarget, 'beforePts')} />
      </label>
      <label>{t('After paragraph (pt)')}
        <input type="number" min="0" step="0.01" value={after ?? ''} placeholder={t(after === undefined ? 'Mixed' : 'Inherit')} onchange={e => paragraphSpacing(e.currentTarget, 'afterPts')} />
      </label>
    </div>
  </section>
{/if}

<style>
  .paragraphs { padding: 12px; border-bottom: 1px solid var(--ok-border); display: grid; gap: 8px; }
  .grid { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); gap: 8px; }
  label { display: flex; flex-direction: column; gap: 4px; font-size: 11px; color: var(--ok-text-2); min-width: 0; }
  input, select { width: 100%; min-width: 0; box-sizing: border-box; font: inherit; padding: 5px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius); background: var(--ok-panel); color: var(--ok-text); }
</style>
