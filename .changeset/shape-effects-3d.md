---
'@office-kit/pptx': minor
---

Shapes can now carry every effect PowerPoint's Format Shape pane edits, one at a time:

- New `setShapeInnerShadow`, `setShapeReflection` and `setShapeSoftEdge`. `setShapeShadow` and `setShapeGlow` also accept `null`, which removes only that effect (other effects stay; a shape whose style references a theme effect keeps an empty effect list so the theme effect does not come back).
- `ShadowOptions` gains `scaleX`, `scaleY`, `skewX` and `skewY` (`sx`, `sy`, `kx`, `ky`), used by PowerPoint's shadow Size field and its perspective shadows. Shape and run shadows read them back.
- New `setShape3D` / `getShape3D` write and read a shape's own `<a:scene3d>` and `<a:sp3d>` (bevels, depth, contour, material, lighting, camera rotation, perspective, distance from ground). The shared 3-D vocabulary (`Text3D`, also exported as `Shape3D`) gains `bevelBottom`, `extrusionColor`, `contourWidthEmu`, `distanceFromGroundEmu`, `scene.cameraRotation` and `scene.fieldOfViewDeg`. Because these are now modelled, `setShapeText3D` replaces them like its other fields: a field left out is removed rather than kept from the deck, and `getShapeText3D` returns them.
