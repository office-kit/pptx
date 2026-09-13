<script lang="ts">
  // Renders one editable control for a capability parameter, chosen by its
  // `kind`. This single component powers both the auto-generated command
  // dialog and any ribbon control that wants a quick inline editor — so every
  // one of the 147 capabilities has a usable input even before a bespoke UI is
  // written for it. `object` params fall back to a structured JSON editor,
  // which keeps the long tail reachable without hand-authoring every schema.
  import { cm, emu, inches, pt } from '@office-kit/pptx';
  import type { ParamSpec } from '../manifest/types.ts';
  import ParamField from './ParamField.svelte';

  interface Props {
    spec: ParamSpec;
    value: unknown;
    onchange: (value: unknown) => void;
  }
  let { spec, value, onchange }: Props = $props();

  const label = $derived(spec.label ?? spec.name);

  // EMU unit handling ------------------------------------------------------
  const UNITS = ['in', 'cm', 'pt', 'emu'] as const;
  type Unit = (typeof UNITS)[number];
  let unit = $state<Unit>('in');
  function toEmu(n: number, u: Unit): number {
    switch (u) {
      case 'in': return inches(n) as unknown as number;
      case 'cm': return cm(n) as unknown as number;
      case 'pt': return pt(n) as unknown as number;
      case 'emu': return emu(n) as unknown as number;
    }
  }
  function fromEmu(e: number, u: Unit): number {
    const perUnit = toEmu(1, u);
    return perUnit ? Math.round((e / perUnit) * 1000) / 1000 : e;
  }

  // Color: the API accepts hex strings ('RRGGBB' / '#RRGGBB'). Keep a text
  // field for scheme refs, plus a swatch for quick picking.
  function normHex(v: string): string {
    return v.startsWith('#') ? v.slice(1) : v;
  }
  function displayHex(v: unknown): string {
    const s = typeof v === 'string' ? v : '';
    const h = normHex(s);
    return /^[0-9a-fA-F]{6}$/.test(h) ? `#${h}` : '#000000';
  }
</script>

<div class="ok-field">
  <label for={`p-${spec.name}`}>
    {label}
    {#if spec.optional}<span class="opt">(optional)</span>{/if}
  </label>

  {#if spec.kind === 'string'}
    <input
      id={`p-${spec.name}`}
      class="ok-input"
      type="text"
      value={(value as string) ?? ''}
      oninput={(e) => onchange(e.currentTarget.value)}
    />
  {:else if spec.kind === 'number' || spec.kind === 'index'}
    <input
      id={`p-${spec.name}`}
      class="ok-input"
      type="number"
      step={spec.kind === 'index' ? 1 : 'any'}
      value={(value as number) ?? ''}
      oninput={(e) => onchange(e.currentTarget.value === '' ? undefined : Number(e.currentTarget.value))}
    />
  {:else if spec.kind === 'emu'}
    <div class="row">
      <input
        id={`p-${spec.name}`}
        class="ok-input"
        type="number"
        step="any"
        value={value == null ? '' : fromEmu(value as number, unit)}
        oninput={(e) =>
          onchange(e.currentTarget.value === '' ? undefined : toEmu(Number(e.currentTarget.value), unit))}
      />
      <select class="ok-select" bind:value={unit} aria-label="unit">
        {#each UNITS as u (u)}<option value={u}>{u}</option>{/each}
      </select>
    </div>
  {:else if spec.kind === 'color'}
    <div class="row">
      <input
        type="color"
        class="swatch"
        aria-label={`${label} swatch`}
        value={displayHex(value)}
        oninput={(e) => onchange(normHex(e.currentTarget.value))}
      />
      <input
        id={`p-${spec.name}`}
        class="ok-input"
        type="text"
        placeholder="RRGGBB or scheme:accent1"
        value={(value as string) ?? ''}
        oninput={(e) => onchange(e.currentTarget.value)}
      />
    </div>
  {:else if spec.kind === 'boolean'}
    <label class="chk">
      <input
        id={`p-${spec.name}`}
        type="checkbox"
        checked={Boolean(value)}
        onchange={(e) => onchange(e.currentTarget.checked)}
      />
      <span>{label}</span>
    </label>
  {:else if spec.kind === 'enum'}
    <select
      id={`p-${spec.name}`}
      class="ok-select"
      value={(value as string) ?? ''}
      onchange={(e) => onchange(e.currentTarget.value)}
    >
      <option value="" disabled>Choose…</option>
      {#each spec.enumValues ?? [] as opt (opt)}
        <option value={opt}>{opt}</option>
      {/each}
    </select>
  {:else if spec.kind === 'object' && spec.fields}
    <!-- Nested object: render each sub-field and assemble into an object. -->
    {@const obj = (value ?? {}) as Record<string, unknown>}
    <div class="object-group">
      {#each spec.fields as f (f.name)}
        <ParamField
          spec={f}
          value={obj[f.name]}
          onchange={(v) => {
            const next = { ...obj };
            if (v === undefined) delete next[f.name];
            else next[f.name] = v;
            onchange(next);
          }}
        />
      {/each}
    </div>
  {:else if spec.kind === 'array' && spec.item}
    {@const arr = (Array.isArray(value) ? value : []) as unknown[]}
    <div class="array-group">
      {#each arr as el, i (i)}
        <div class="array-item">
          <ParamField
            spec={{ ...spec.item, name: `${spec.name}[${i}]`, label: `#${i + 1}` }}
            value={el}
            onchange={(v) => {
              const next = arr.slice();
              next[i] = v;
              onchange(next);
            }}
          />
          <button type="button" class="ok-btn tiny" title="Remove" onclick={() => onchange(arr.filter((_, j) => j !== i))}>✕</button>
        </div>
      {/each}
      <button type="button" class="ok-btn tiny add" onclick={() => onchange([...arr, undefined])}>＋ Add</button>
    </div>
  {:else}
    <!-- object / unknown without a schema → structured JSON, still reachable -->
    <textarea
      id={`p-${spec.name}`}
      class="ok-input json"
      rows="3"
      placeholder={spec.type}
      value={value == null ? '' : JSON.stringify(value, null, 0)}
      oninput={(e) => {
        const raw = e.currentTarget.value.trim();
        if (raw === '') return onchange(undefined);
        try {
          onchange(JSON.parse(raw));
          e.currentTarget.setCustomValidity('');
        } catch {
          e.currentTarget.setCustomValidity('Invalid JSON');
        }
      }}
    ></textarea>
    <span class="type-hint">{spec.type}</span>
  {/if}
</div>

<style>
  .row {
    display: flex;
    gap: 6px;
    align-items: center;
  }
  .row > .ok-input {
    flex: 1;
  }
  .row > .ok-select {
    width: 64px;
    flex: none;
  }
  .swatch {
    width: 34px;
    height: 26px;
    padding: 0;
    border: 1px solid var(--ok-border-strong);
    border-radius: var(--ok-radius);
    background: none;
    cursor: pointer;
    flex: none;
  }
  .chk {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 13px;
  }
  .opt {
    color: var(--ok-text-3);
    font-size: 10px;
  }
  .json {
    font-family: var(--ok-mono);
    font-size: 11px;
    resize: vertical;
  }
  .type-hint {
    font-family: var(--ok-mono);
    font-size: 10px;
    color: var(--ok-text-3);
    word-break: break-all;
  }
  .object-group,
  .array-group {
    border-left: 2px solid var(--ok-border);
    padding-left: 10px;
    margin-top: 2px;
    display: flex;
    flex-direction: column;
  }
  .array-item {
    display: flex;
    gap: 6px;
    align-items: flex-start;
  }
  .array-item > :global(.ok-field) {
    flex: 1;
  }
  .ok-btn.tiny {
    padding: 2px 6px;
    font-size: 11px;
    border: 1px solid var(--ok-border);
  }
  .ok-btn.tiny.add {
    align-self: flex-start;
    margin-top: 4px;
  }
</style>
