<script lang="ts">
  import { getShapeStrokeDash, getShapeStrokeCap, getShapeStrokeJoin, getShapeStrokeCompound } from '@office-kit/pptx';
  import ArrowStyleFields from './ArrowStyleFields.svelte';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const fields = [
    { label: 'Compound type', read: getShapeStrokeCompound, command: 'setShapeStrokeCompound', parameter: 'cmpd', options: [['sng', 'Single'], ['dbl', 'Double'], ['thickThin', 'Thick Thin'], ['thinThick', 'Thin Thick'], ['tri', 'Triple']] },
    { label: 'Dash type', accessibleLabel: 'Outline style', read: getShapeStrokeDash, command: 'setShapeStrokeDash', parameter: 'dash', options: [
    ['solid', 'Solid line'], ['dot', 'Dotted line'], ['dash', 'Dashed line'],
    ['lgDash', 'Long dashed line'], ['dashDot', 'Dash-dot line'],
    ['lgDashDot', 'Long dash-dot line'], ['lgDashDotDot', 'Long dash-dot-dot line'],
    ['sysDash', 'System dashed line'], ['sysDot', 'System dotted line'],
    ['sysDashDot', 'System dash-dot line'], ['sysDashDotDot', 'System dash-dot-dot line'],
  ] },
    { label: 'Cap type', read: getShapeStrokeCap, command: 'setShapeStrokeCap', parameter: 'cap', options: [['flat', 'Flat'], ['rnd', 'Round'], ['sq', 'Square']] },
    { label: 'Join type', read: getShapeStrokeJoin, command: 'setShapeStrokeJoin', parameter: 'join', options: [['round', 'Round'], ['bevel', 'Bevel'], ['miter', 'Miter']] },
  ] as const;
  const values = $derived.by(() => {
    editor.doc.version;
    const shapes = editor.selectedShapes();
    return fields.map(field => {
      const values = new Set(shapes.map(shape => field.read(shape) ?? 'inherit'));
      return values.size === 1 ? [...values][0]! : 'mixed';
    });
  });
</script>

{#each fields as field, index}
  <label>
    <span>{t(field.label)}</span>
    <select class="ok-input" aria-label={t('accessibleLabel' in field ? field.accessibleLabel : field.label)} value={values[index]}
      onchange={event => editor.invoke(field.command, { [field.parameter]: event.currentTarget.value })}>
      {#if values[index] === 'inherit' || values[index] === 'mixed'}
        <option value={values[index]} disabled>{t(values[index] === 'mixed' ? 'Mixed' : 'Inherited')}</option>
      {/if}
      {#each field.options as [value, label]}<option {value}>{t(label)}</option>{/each}
    </select>
  </label>
{/each}

<ArrowStyleFields />

<style>
  label { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
  select { max-width: 145px; font-size: inherit; }
</style>
