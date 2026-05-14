// Node-specific entry point. P1+ will add fs-backed convenience methods like
// `Presentation.loadFile(path)` and `Presentation.prototype.saveTo(path)` here.
// For now it just re-exports the public API.

export * from './api/index.ts';
