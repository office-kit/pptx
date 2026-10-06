---
'@office-kit/pptx': minor
---

Text formats can now carry PowerPoint's theme tints and WordArt bevels.

- `TextFormat.colorTransforms`, and `colorTransforms` on `outline`, `shadow`, `innerShadow` and `glow`, write `<a:lumMod>`, `<a:lumOff>`, `<a:tint>`, ... on the run's colors — the same field gradient stops already take — so "Accent 2, Lighter 60%" stays a theme color instead of a fixed RGB. `getShapeRunFormat` reads them back beside the unresolved color and `toWritableTextFormat` carries them. The shared `ShadowOptions`, `InnerShadowOptions` and `GlowOptions` gain the field, so `setShapeShadow` and `setShapeGlow` accept it too.
- New `setShapeText3D(shape, value | null)` / `getShapeText3D(shape)` write and read the text body's 3-D (`<a:scene3d>` camera and light rig, `<a:sp3d>` top bevel, extrusion height, material and contour color), which is where PowerPoint puts its WordArt bevels. Presets are typed by the schema's enums (`CameraPreset`, `LightRigType`, `LightRigDirection`, `BevelPreset`, `PresetMaterial`).
- `setShapeTextAutoFit` now places its element ahead of any `<a:scene3d>` / `<a:sp3d>`, as the schema requires.
- `getShapeRunFormat` reports a run effect's scheme color as its token (it returned an empty string without a theme), and a text gradient stop without alpha no longer reads back as `opacity: 1`, which wrote an `<a:alpha val="100000"/>` on the way back.
