<script lang="ts">
  import { getShapeKind, type Color, type ColorTransform, type ReflectionOptions, type SlideShapeData, type Text3D } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { getLocale, t } from '../i18n/i18n.svelte.ts';
  import ColorPicker from '../ui/ColorPicker.svelte';
  import PaneSection from './PaneSection.svelte';
  import PresetMenu from './PresetMenu.svelte';
  import SliderField from './SliderField.svelte';
  import {
    BEVEL_PRESETS, DEFAULT_BEVEL_EMU, DEFAULT_SCENE, EMU_PER_POINT, GLOW_PRESETS, LIGHTING_PRESETS, MATERIAL_PRESETS,
    REFLECTION_PRESETS, ROTATION_PRESETS, SHADOW_PRESETS, SOFT_EDGE_PRESETS, writable3D,
    type EffectTarget, type EffectsState, type GlowValue, type ShadowValue,
  } from './effects-model.ts';

  // PowerPoint's Effects category, in its order: Shadow, Reflection, Glow,
  // Soft Edges, 3-D Format, 3-D Rotation. The same sections serve Shape
  // Options ▸ Effects and Text Options ▸ Text Effects through `target`.
  let { target }: { target: EffectTarget } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  const prefix = $derived(target.kind === 'text' ? 'text-' : '');
  const shapes = $derived.by(() => {
    doc.version;
    return editor.selectedShapes().filter((shape) => target.kind === 'text'
      ? getShapeKind(shape) === 'shape'
      : ['shape', 'connector', 'picture'].includes(getShapeKind(shape)));
  });
  const disabled = $derived(!shapes.length || editor.selectionLocked());
  const states = $derived.by(() => { doc.version; return shapes.map((shape) => target.read(doc.pres, shape)); });
  // Japanese PowerPoint labels the galleries 標準スタイル, while other
  // "Presets" in the editor (video corrections) read プリセット.
  const presetsLabel = $derived(getLocale() === 'ja' ? '標準スタイル' : 'Presets');

  function common<T>(read: (state: EffectsState) => T | undefined): T | undefined {
    const values = states.map(read);
    return values.length && values.every((value) => value === values[0]) ? values[0] : undefined;
  }
  function apply(label: string, change: (shape: SlideShapeData, state: EffectsState) => void) {
    if (disabled) return;
    const items = shapes.map((shape, index) => ({ shape, state: states[index]! }));
    doc.transact(t(label), () => { for (const { shape, state } of items) change(shape, state); });
  }
  const pt = (emu: number | undefined) => emu === undefined ? undefined : emu / EMU_PER_POINT;
  const percent = (fraction: number | undefined) => fraction === undefined ? undefined : fraction * 100;

  // --- Shadow
  const DEFAULT_SHADOW = SHADOW_PRESETS[0]!.items[0]!.value;
  const near = (a: number | undefined, b: number | undefined, tolerance = 0.01) => Math.abs((a ?? 0) - (b ?? 0)) <= tolerance;
  const shadowPreset = $derived(common((state) => {
    const value = state.shadow;
    if (!value) return undefined;
    return SHADOW_PRESETS.flatMap((group) => group.items).find(({ value: preset }) =>
      preset.kind === value.kind && near(preset.angleDeg, value.angleDeg, 0.5) && preset.blurEmu === value.blurEmu &&
      preset.offsetEmu === value.offsetEmu && near(preset.opacity, value.opacity) && near(preset.scaleY ?? 1, value.scaleY ?? 1) &&
      near(preset.skewX, value.skewX, 0.5))?.label;
  }));
  const hasShadow = $derived(states.some((state) => state.shadow));
  function shadow(label: string, update: Partial<ShadowValue> | null) {
    apply(label, (shape, state) => target.shadow(shape, update && { ...(state.shadow ?? DEFAULT_SHADOW), ...update }));
  }
  const innerShadow = $derived(states.some((state) => state.shadow?.kind === 'inner'));

  // --- Reflection
  const DEFAULT_REFLECTION = REFLECTION_PRESETS[0]!.value;
  const reflectionPreset = $derived(common((state) => {
    const value = state.reflection;
    if (!value) return undefined;
    return REFLECTION_PRESETS.find(({ value: preset }) => near(preset.endPosition, value.endPosition) &&
      (preset.offsetEmu ?? 0) === (value.offsetEmu ?? 0) && near(preset.startOpacity, value.startOpacity))?.label;
  }));
  function reflection(label: string, update: Partial<ReflectionOptions> | null) {
    apply(label, (shape, state) => target.reflection(shape, update && { ...(state.reflection ?? DEFAULT_REFLECTION), ...update }));
  }

  // --- Glow
  const DEFAULT_GLOW = GLOW_PRESETS[0]!.value;
  const glowPreset = $derived(common((state) => {
    const value = state.glow;
    if (!value) return undefined;
    return GLOW_PRESETS.find(({ value: preset }) => preset.radiusEmu === value.radiusEmu && preset.color === value.color && near(preset.opacity, value.opacity))?.label;
  }));
  function glow(label: string, update: Partial<GlowValue> | null) {
    apply(label, (shape, state) => target.glow(shape, update && { ...(state.glow ?? DEFAULT_GLOW), ...update }));
  }

  // --- Soft Edges
  const softEdgePreset = $derived(common((state) => SOFT_EDGE_PRESETS.find((preset) => preset.value === state.softEdgeEmu)?.label));
  const softEdgeReason = $derived(target.softEdge ? undefined : t('Soft edges on text are not supported by the library yet.'));
  function softEdge(radiusEmu: number | null) {
    const write = target.softEdge;
    if (write) apply('Soft Edges', (shape) => write(shape, radiusEmu));
  }

  // --- 3-D
  function threeD(label: string, update: (value: Text3D) => Text3D) {
    apply(label, (shape, state) => {
      const next = update(writable3D(state.threeD));
      const { scene, ...shape3D } = next;
      const hasShape3D = Object.keys(shape3D).length > 0;
      const plainScene = !scene || (scene.camera === DEFAULT_SCENE.camera && !scene.cameraRotation && scene.fieldOfViewDeg === undefined &&
        scene.lightRig.type === DEFAULT_SCENE.lightRig.type && scene.lightRig.direction === DEFAULT_SCENE.lightRig.direction && !scene.lightRig.rotation);
      // Nothing left to draw in 3-D: drop both elements rather than keep a default scene.
      if (!hasShape3D && plainScene) { target.threeD(shape, null); return; }
      target.threeD(shape, { ...next, scene: scene ?? DEFAULT_SCENE });
    });
  }
  const bevel = (which: 'bevelTop' | 'bevelBottom') => ({
    preset: common((state) => state.threeD?.[which] ? (state.threeD[which]!.preset ?? 'circle') : undefined),
    width: common((state) => pt(state.threeD?.[which] ? (state.threeD[which]!.widthEmu ?? DEFAULT_BEVEL_EMU) : undefined)),
    height: common((state) => pt(state.threeD?.[which] ? (state.threeD[which]!.heightEmu ?? DEFAULT_BEVEL_EMU) : undefined)),
  });
  const bevels = $derived({ bevelTop: bevel('bevelTop'), bevelBottom: bevel('bevelBottom') });
  function setBevel(which: 'bevelTop' | 'bevelBottom', update: Partial<NonNullable<Text3D['bevelTop']>> | null) {
    threeD(which === 'bevelTop' ? 'Top bevel' : 'Bottom bevel', (value) => {
      const { [which]: current, ...rest } = value;
      return update ? { ...rest, [which]: { widthEmu: DEFAULT_BEVEL_EMU, heightEmu: DEFAULT_BEVEL_EMU, ...current, ...update } } : rest;
    });
  }
  const camera = $derived(common((state) => state.threeD?.scene?.camera));
  const rotationPreset = $derived(ROTATION_PRESETS.flatMap((group) => group.items).find((item) => item.value === camera)?.label);
  const rotation = $derived({
    x: common((state) => state.threeD?.scene ? (state.threeD.scene.cameraRotation?.longitudeDeg ?? 0) : 0),
    y: common((state) => state.threeD?.scene ? (state.threeD.scene.cameraRotation?.latitudeDeg ?? 0) : 0),
    z: common((state) => state.threeD?.scene ? (state.threeD.scene.cameraRotation?.revolutionDeg ?? 0) : 0),
  });
  function rotate(axis: 'longitudeDeg' | 'latitudeDeg' | 'revolutionDeg', degrees: number) {
    const label = { longitudeDeg: 'X Rotation', latitudeDeg: 'Y Rotation', revolutionDeg: 'Z Rotation' }[axis];
    threeD(label, (value) => {
      const scene = value.scene ?? DEFAULT_SCENE;
      const cameraRotation = { latitudeDeg: 0, longitudeDeg: 0, revolutionDeg: 0, ...scene.cameraRotation, [axis]: degrees % 360 };
      return { ...value, scene: { ...scene, cameraRotation } };
    });
  }
  const BEVEL_ROWS = [['bevelTop', 'Top bevel'], ['bevelBottom', 'Bottom bevel']] as const;
  function withColor(value: Text3D, field: 'extrusionColor' | 'contourColor', color: Color, transforms: readonly ColorTransform[] | undefined): Text3D {
    const key = field === 'extrusionColor' ? 'extrusionColorTransforms' : 'contourColorTransforms';
    const { [key]: _previous, ...rest } = value;
    return { ...rest, [field]: color, ...(transforms?.length ? { [key]: transforms } : {}) };
  }
  // Native nudge buttons beside each rotation box; the steps are assumed.
  const ROTATION_ROWS = [
    { axis: 'x', field: 'longitudeDeg', label: 'X Rotation', nudges: [['Left', '左へ', -10], ['Right', '右へ', 10]] },
    { axis: 'y', field: 'latitudeDeg', label: 'Y Rotation', nudges: [['Up', '上へ', -10], ['Down', '下へ', 10]] },
    { axis: 'z', field: 'revolutionDeg', label: 'Z Rotation', nudges: [['Clockwise', '時計回り', 10], ['Counter-clockwise', '反時計回り', -10]] },
  ] as const;
  const FOV_NUDGES = [['Narrow field of view', '狭角', -5], ['Widen field of view', '広角', 5]] as const;
const nudgeName = (english: string, japanese: string) => (getLocale() === 'ja' ? japanese : english);
  const fieldOfView = $derived(common((state) => state.threeD?.scene?.fieldOfViewDeg ?? 45));
  function perspective(fieldOfViewDeg: number) {
    threeD('Perspective', (value) => ({ ...value, scene: { ...(value.scene ?? DEFAULT_SCENE), fieldOfViewDeg } }));
  }
  const perspectiveCamera = $derived(camera?.startsWith('perspective') ?? false);
</script>

{#snippet colorRow(label: string, value: string | undefined, choose: (color: Color, transforms?: readonly ColorTransform[]) => void, enabled = true)}
  <div class="pane-row"><span>{t('Color')}</span><span class="pane-color"><ColorPicker label={label} {value} resolvedColor={value && /^#[0-9a-f]{6}$/i.test(value) ? value : undefined} disabled={disabled || !enabled} {choose} /></span></div>
{/snippet}
{#snippet shadowTile(value: ShadowValue | null)}
  {@const radians = ((value?.angleDeg ?? 0) * Math.PI) / 180}
  {@const offset = value && value.offsetEmu ? 4 : 0}
  <span class="tile" style:box-shadow={value ? `${value.kind === 'inner' ? 'inset ' : ''}${(Math.cos(radians) * offset).toFixed(1)}px ${(Math.sin(radians) * offset).toFixed(1)}px ${value.blurEmu ? 4 : 0}px rgba(0,0,0,${Math.min(0.8, value.opacity + 0.2)})` : 'none'}></span>
{/snippet}
{#snippet reflectionTile(value: ReflectionOptions | null)}
  <span class="tile reflect" style:--reflect-size="{Math.round((value?.endPosition ?? 0) * 100)}%" style:--reflect-gap="{value?.offsetEmu ? Math.round(value.offsetEmu / EMU_PER_POINT / 2) : 0}px" class:on={!!value}></span>
{/snippet}
{#snippet glowTile(value: GlowValue | null)}
  <span class="tile" style:box-shadow={value ? `0 0 ${Math.round(value.radiusEmu / EMU_PER_POINT / 2)}px ${Math.round(value.radiusEmu / EMU_PER_POINT / 4)}px var(--glow-${value.color}, var(--ok-accent))` : 'none'}></span>
{/snippet}
{#snippet softTile(value: number | null)}
  <span class="tile" style:filter={value ? `blur(${Math.min(6, value / EMU_PER_POINT / 4)}px)` : undefined}></span>
{/snippet}

<div class="effects" data-effects-target={target.kind}>
  <PaneSection id="{prefix}shadow" label={t('Shadow')}>
    <div class="pane-row"><span>{presetsLabel}</span>
      <PresetMenu label={presetsLabel} groups={SHADOW_PRESETS} selected={shadowPreset} {disabled} columns={3} tile={shadowTile}
        none={{ label: 'No Shadow', selected: !hasShadow, choose: () => shadow('Shadow', null) }}
        choose={(value) => apply('Shadow', (shape) => target.shadow(shape, value))} />
    </div>
    {@render colorRow(t('Shadow Color'), common((state) => state.shadow ? String(state.shadow.color) : undefined), (color, colorTransforms) => shadow('Shadow Color', { color, colorTransforms: colorTransforms ?? [] }))}
    <SliderField label={t('Transparency')} name={t('Shadow transparency')} value={common((state) => state.shadow ? Math.round((1 - state.shadow.opacity) * 100) : undefined)} min={0} max={100} unit="%" {disabled} apply={(value) => shadow('Shadow', { opacity: 1 - value / 100 })} />
    <SliderField label={t('Size')} name={t('Shadow size')} value={innerShadow ? undefined : common((state) => state.shadow ? percent(state.shadow.scaleX ?? 1) : undefined)} min={1} max={200} unit="%" disabled={disabled || innerShadow} apply={(value) => shadow('Shadow', { scaleX: value / 100, scaleY: value / 100 })} />
    <SliderField label={t('Blur')} name={t('Shadow blur')} value={common((state) => pt(state.shadow?.blurEmu))} min={0} max={100} unit="pt" {disabled} apply={(value) => shadow('Shadow', { blurEmu: Math.round(value * EMU_PER_POINT) })} />
    <SliderField label={t('Angle')} name={t('Shadow angle')} value={common((state) => state.shadow?.angleDeg)} min={0} max={359.9} unit="°" {disabled} apply={(value) => shadow('Shadow', { angleDeg: value })} />
    <SliderField label={t('Distance')} name={t('Shadow distance')} value={common((state) => pt(state.shadow?.offsetEmu))} min={0} max={200} unit="pt" {disabled} apply={(value) => shadow('Shadow', { offsetEmu: Math.round(value * EMU_PER_POINT) })} />
  </PaneSection>

  <PaneSection id="{prefix}reflection" label={t('Reflection')}>
    <div class="pane-row"><span>{presetsLabel}</span>
      <PresetMenu label={presetsLabel} groups={[{ heading: 'Reflection Variations', ja: '反射の種類', items: REFLECTION_PRESETS }]} selected={reflectionPreset} {disabled} columns={3} tile={reflectionTile}
        none={{ label: 'No Reflection', selected: !states.some((state) => state.reflection), choose: () => reflection('Reflection', null) }}
        choose={(value) => apply('Reflection', (shape) => target.reflection(shape, value))} />
    </div>
    <SliderField label={t('Transparency')} name={t('Reflection transparency')} value={common((state) => state.reflection ? Math.round((1 - (state.reflection.startOpacity ?? 1)) * 100) : undefined)} min={0} max={100} unit="%" {disabled} apply={(value) => reflection('Reflection', { startOpacity: 1 - value / 100 })} />
    <SliderField label={t('Size')} name={t('Reflection size')} value={common((state) => state.reflection ? percent(state.reflection.endPosition ?? 1) : undefined)} min={1} max={100} unit="%" {disabled} apply={(value) => reflection('Reflection', { endPosition: value / 100 })} />
    <SliderField label={t('Blur')} name={t('Reflection blur')} value={common((state) => state.reflection ? pt(state.reflection.blurEmu ?? 0) : undefined)} min={0} max={100} unit="pt" {disabled} apply={(value) => reflection('Reflection', { blurEmu: Math.round(value * EMU_PER_POINT) })} />
    <SliderField label={t('Distance')} name={t('Reflection distance')} value={common((state) => state.reflection ? pt(state.reflection.offsetEmu ?? 0) : undefined)} min={0} max={100} unit="pt" {disabled} apply={(value) => reflection('Reflection', { offsetEmu: Math.round(value * EMU_PER_POINT) })} />
  </PaneSection>

  <PaneSection id="{prefix}glow" label={t('Glow')}>
    <div class="pane-row"><span>{presetsLabel}</span>
      <PresetMenu label={presetsLabel} groups={[{ heading: 'Glow Variations', ja: '光彩の種類', items: GLOW_PRESETS }]} selected={glowPreset} {disabled} columns={6} tile={glowTile}
        none={{ label: 'No Glow', selected: !states.some((state) => state.glow), choose: () => glow('Glow', null) }}
        choose={(value) => apply('Glow', (shape) => target.glow(shape, value))} />
    </div>
    {@render colorRow(t('Glow Color'), common((state) => state.glow ? String(state.glow.color) : undefined), (color, colorTransforms) => glow('Glow Color', { color, colorTransforms: colorTransforms ?? [] }))}
    <SliderField label={t('Size')} name={t('Glow size')} value={common((state) => pt(state.glow?.radiusEmu))} min={0} max={150} unit="pt" {disabled} apply={(value) => glow('Glow', { radiusEmu: Math.round(value * EMU_PER_POINT) })} />
    <SliderField label={t('Transparency')} name={t('Glow transparency')} value={common((state) => state.glow ? Math.round((1 - state.glow.opacity) * 100) : undefined)} min={0} max={100} unit="%" {disabled} apply={(value) => glow('Glow', { opacity: 1 - value / 100 })} />
  </PaneSection>

  <PaneSection id="{prefix}softEdges" label={t('Soft Edges')}>
    <div class="pane-row" title={softEdgeReason}><span>{presetsLabel}</span>
      <PresetMenu label={presetsLabel} title={softEdgeReason} groups={[{ heading: 'Soft Edge Variations', ja: 'ぼかしの種類', items: SOFT_EDGE_PRESETS }]} selected={softEdgePreset} disabled={disabled || !!softEdgeReason} columns={3} tileSize={44} tile={softTile}
        none={{ label: 'No Soft Edges', selected: !states.some((state) => state.softEdgeEmu), choose: () => softEdge(null) }}
        choose={(value) => softEdge(value)} />
    </div>
    <SliderField label={t('Size')} name={t('Soft edge size')} title={softEdgeReason} value={common((state) => pt(state.softEdgeEmu ?? undefined))} min={0} max={100} unit="pt" disabled={disabled || !!softEdgeReason} apply={(value) => softEdge(value ? Math.round(value * EMU_PER_POINT) : null)} />
  </PaneSection>

  <PaneSection id="{prefix}3dFormat" label={t('3-D Format')}>
    <!-- Native: a caption row, then a 59 × 60 pt gallery button at the left
         with its Width / Height (or Size / Angle) boxes beside it. -->
    {#each BEVEL_ROWS as [which, label]}
      <div class="caption">{t(label)}</div>
      <div class="gallery-row">
        <PresetMenu label={t(label)} size="large" tileSize={52} groups={[{ heading: 'Bevel', ja: '面取り', items: BEVEL_PRESETS }]} selected={BEVEL_PRESETS.find((preset) => preset.value === bevels[which].preset)?.label} {disabled} columns={4}
          none={{ label: 'No Bevel', selected: !states.some((state) => state.threeD?.[which]), choose: () => setBevel(which, null) }}
          choose={(preset) => setBevel(which, { preset })} />
        <div class="side">
          <SliderField label={t('Width')} name={t(`${label} width`)} slider={false} value={bevels[which].width ?? 0} min={0} max={1584} unit="pt" {disabled} apply={(value) => setBevel(which, { widthEmu: Math.round(value * EMU_PER_POINT) })} />
          <SliderField label={t('Height')} name={t(`${label} height`)} slider={false} value={bevels[which].height ?? 0} min={0} max={1584} unit="pt" {disabled} apply={(value) => setBevel(which, { heightEmu: Math.round(value * EMU_PER_POINT) })} />
        </div>
      </div>
    {/each}
    <div class="caption">{t('Depth')}</div>
    <div class="gallery-row">
      <span class="pane-color"><ColorPicker label={t('Depth')} value={common((state) => state.threeD?.extrusionColor)} {disabled} choose={(color, colorTransforms) => threeD('Depth', (value) => withColor(value, 'extrusionColor', color, colorTransforms))} /></span>
      <div class="side"><SliderField label={t('Size')} name={t('Depth size')} slider={false} value={common((state) => pt(state.threeD?.extrusionHeightEmu ?? 0))} min={0} max={1584} unit="pt" {disabled} apply={(size) => threeD('Depth', (value) => ({ ...value, extrusionHeightEmu: Math.round(size * EMU_PER_POINT) }))} /></div>
    </div>
    <div class="caption">{t('Contour')}</div>
    <div class="gallery-row">
      <span class="pane-color"><ColorPicker label={t('Contour')} value={common((state) => state.threeD?.contourColor)} {disabled} choose={(color, colorTransforms) => threeD('Contour', (value) => withColor(value, 'contourColor', color, colorTransforms))} /></span>
      <div class="side"><SliderField label={t('Size')} name={t('Contour size')} slider={false} value={common((state) => pt(state.threeD?.contourWidthEmu ?? 0))} min={0} max={1584} unit="pt" {disabled} apply={(size) => threeD('Contour', (value) => ({ ...value, contourWidthEmu: Math.round(size * EMU_PER_POINT) }))} /></div>
    </div>
    <div class="caption">{t('Material')}</div>
    <div class="gallery-row">
      <PresetMenu label={t('Material')} size="large" tileSize={52} groups={MATERIAL_PRESETS} selected={MATERIAL_PRESETS.flatMap((group) => group.items).find((item) => item.value === common((state) => state.threeD?.material))?.label} {disabled} columns={4}
        choose={(material) => threeD('Material', (value) => ({ ...value, material }))} />
    </div>
    <div class="caption">{t('Lighting')}</div>
    <div class="gallery-row">
      <PresetMenu label={t('Lighting')} size="large" tileSize={52} groups={LIGHTING_PRESETS} selected={LIGHTING_PRESETS.flatMap((group) => group.items).find((item) => item.value === common((state) => state.threeD?.scene?.lightRig.type))?.label} {disabled} columns={4}
        choose={(type) => threeD('Lighting', (value) => { const scene = value.scene ?? DEFAULT_SCENE; return { ...value, scene: { ...scene, lightRig: { ...scene.lightRig, type } } }; })} />
      <div class="side"><SliderField label={t('Angle')} name={t('Lighting angle')} slider={false} value={common((state) => state.threeD?.scene?.lightRig.rotation?.revolutionDeg ?? 0)} min={0} max={359.9} unit="°" {disabled}
        apply={(degrees) => threeD('Lighting', (value) => { const scene = value.scene ?? DEFAULT_SCENE; return { ...value, scene: { ...scene, lightRig: { ...scene.lightRig, rotation: { latitudeDeg: 0, longitudeDeg: 0, ...scene.lightRig.rotation, revolutionDeg: degrees } } } }; })} /></div>
    </div>
    <button class="ok-btn reset" aria-label={t('Reset 3-D Format')} {disabled} onclick={() => threeD('Reset 3-D Format', (value) => (value.scene ? { scene: { ...value.scene, lightRig: DEFAULT_SCENE.lightRig } } : {}))}>{t('Reset')}</button>
  </PaneSection>

  <PaneSection id="{prefix}3dRotation" label={t('3-D Rotation')}>
    <div class="presets-row"><span>{presetsLabel}</span>
      <PresetMenu label={presetsLabel} size="large" tileSize={52} groups={ROTATION_PRESETS} selected={rotationPreset} {disabled} columns={4}
        none={{ label: 'No Rotation', selected: !camera || camera === 'orthographicFront', choose: () => threeD('3-D Rotation', (value) => (value.scene ? { ...value, scene: { camera: 'orthographicFront', lightRig: value.scene.lightRig } } : value)) }}
        choose={(preset) => threeD('3-D Rotation', (value) => ({ ...value, scene: { camera: preset, lightRig: (value.scene ?? DEFAULT_SCENE).lightRig } }))} />
    </div>
    {#each ROTATION_ROWS as row (row.axis)}
      <div class="nudge-row">
        <SliderField label={t(row.label)} slider={false} value={rotation[row.axis]} min={0} max={359.9} unit="°" {disabled} apply={(value) => rotate(row.field, value)} />
        {#each row.nudges as [nudgeLabel, nudgeJa, step] (nudgeLabel)}
          <button class="ok-btn nudge" aria-label={nudgeName(nudgeLabel, nudgeJa)} title={nudgeName(nudgeLabel, nudgeJa)} {disabled} onclick={() => rotate(row.field, ((((rotation[row.axis] ?? 0) + step) % 360) + 360) % 360)}>{step < 0 ? '↺' : '↻'}</button>
        {/each}
      </div>
    {/each}
    <div class="nudge-row">
      <SliderField label={t('Perspective')} slider={false} title={perspectiveCamera ? undefined : t('Perspective applies to perspective presets only.')} value={perspectiveCamera ? fieldOfView : undefined} min={0} max={120} unit="°" disabled={disabled || !perspectiveCamera}
        apply={(fieldOfViewDeg) => perspective(fieldOfViewDeg)} />
      {#each FOV_NUDGES as [nudgeLabel, nudgeJa, step] (nudgeLabel)}
        <button class="ok-btn nudge" aria-label={nudgeName(nudgeLabel, nudgeJa)} title={nudgeName(nudgeLabel, nudgeJa)} disabled={disabled || !perspectiveCamera} onclick={() => perspective(Math.min(120, Math.max(0, (fieldOfView ?? 45) + step)))}>{step < 0 ? '−' : '+'}</button>
      {/each}
    </div>
    <label class="pane-check" title={t('Keep text flat is not supported by the library yet.')}><input type="checkbox" disabled /><span>{t('Keep text flat')}</span></label>
    <SliderField label={t('Distance from ground')} slider={false} value={common((state) => pt(state.threeD?.distanceFromGroundEmu ?? 0))} min={-4000} max={4000} unit="pt" {disabled}
      apply={(points) => threeD('Distance from ground', (value) => ({ ...value, distanceFromGroundEmu: Math.round(points * EMU_PER_POINT) }))} />
    <button class="ok-btn reset" aria-label={t('Reset 3-D Rotation')} {disabled} onclick={() => threeD('Reset 3-D Rotation', (value) => {
      const { distanceFromGroundEmu: _ground, ...rest } = value;
      return rest.scene ? { ...rest, scene: { camera: 'orthographicFront', lightRig: rest.scene.lightRig } } : rest;
    })}>{t('Reset')}</button>
  </PaneSection>
</div>

<style>
  .effects { display: flex; flex-direction: column; }
  /* Native Reset: a 61 × 26 pt button at the left. */
  .reset { align-self: flex-start; min-width: 61px; height: 26px; }
  .caption { min-height: 17px; margin-top: 4px; }
  .gallery-row { display: flex; align-items: flex-start; justify-content: space-between; gap: 8px; min-height: 26px; }
  .gallery-row .side { display: flex; flex-direction: column; gap: 4px; width: 158px; }
  .presets-row { display: flex; align-items: center; justify-content: space-between; min-height: 60px; }
  .nudge-row { display: flex; align-items: center; }
  .nudge-row > :global(.slider-field) { flex: 1; }
  .nudge { width: 32px; height: 26px; padding: 0; flex: none; }
  .tile { display: block; width: 26px; height: 26px; background: var(--ok-bg, #fff); border: 1px solid var(--ok-border-strong); }
  .tile.reflect.on { -webkit-box-reflect: below var(--reflect-gap) linear-gradient(transparent calc(100% - var(--reflect-size)), #0006); }
</style>
