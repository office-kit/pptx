// Tree-shakeable free-function entry points — the canonical public API.
//
// Every operation is a standalone export that operates on the opaque
// `PresentationData` / `SlideData` interfaces defined in
// `_internal-symbols.ts`. Consumers can import only what they use and
// modern bundlers drop the rest.
//
// The implementation is split across `./fn/*.ts` by domain. This file
// is the public barrel; do not add logic here.

export * from './fn/package.ts';
export * from './fn/slides.ts';
export * from './fn/shapes.ts';
export * from './fn/features.ts';
export * from './fn/embedded.ts';
