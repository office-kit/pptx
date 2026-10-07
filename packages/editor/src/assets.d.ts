// A stylesheet's text, for a <style> in the editor's shadow root (see mount.ts).
declare module '*.css?inline' {
  const css: string;
  export default css;
}
