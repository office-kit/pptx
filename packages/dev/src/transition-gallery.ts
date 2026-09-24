export const transitionEffects = [
  ['none', 'None'],
  ['fade', 'Fade'],
  ['push', 'Push'],
  ['wipe', 'Wipe'],
  ['split', 'Split'],
  ['cover', 'Cover'],
  ['cut', 'Cut'],
  ['pull', 'Uncover'],
  ['randomBar', 'Random Bars'],
  ['checker', 'Checkerboard'],
  ['blinds', 'Blinds'],
  ['dissolve', 'Dissolve'],
  ['circle', 'Circle'],
  ['diamond', 'Diamond'],
  ['plus', 'Plus'],
  ['wedge', 'Wedge'],
  ['wheel', 'Wheel'],
  ['wheelReverse', 'Reverse Wheel'],
  ['strips', 'Strips'],
  ['comb', 'Comb'],
  ['zoom', 'Zoom'],
  ['newsflash', 'Newsflash'],
  ['random', 'Random'],
] as const;
export type EffectChoice = {
  effect: string;
  direction?: string;
  orientation?: 'horz' | 'vert';
  thruBlack?: boolean;
  spokes?: number;
};
// ECMA-376 pml.xsd defaults apply when effect attributes are omitted.
export function transitionOptionSelected(
  current: EffectChoice | null | undefined,
  choice: EffectChoice,
) {
  if (!current || current.effect !== choice.effect) return false;
  const defaultDirection = ['split', 'zoom'].includes(current.effect)
    ? 'out'
    : current.effect === 'strips'
      ? 'lu'
      : ['blinds', 'checker', 'comb', 'randomBar'].includes(current.effect)
        ? 'horz'
        : 'l';
  return (
    (choice.direction === undefined ||
      choice.direction === (current.direction ?? defaultDirection)) &&
    (choice.orientation === undefined || choice.orientation === (current.orientation ?? 'horz')) &&
    (choice.thruBlack === undefined || choice.thruBlack === (current.thruBlack ?? false)) &&
    (choice.spokes === undefined || choice.spokes === (current.spokes ?? 4))
  );
}
export function transitionOptions(effect: string): { label: string; value: EffectChoice }[] {
  if (effect === 'wheel' || effect === 'wheelReverse')
    return [1, 2, 3, 4, 8].map((spokes) => ({
      label: spokes + (spokes === 1 ? ' Spoke' : ' Spokes'),
      value: { effect, spokes },
    }));
  if (effect === 'fade' || effect === 'cut')
    return [
      { label: 'Smoothly', value: { effect, thruBlack: false } },
      { label: 'Through Black', value: { effect, thruBlack: true } },
    ];
  if (effect === 'split')
    return (['horz', 'vert'] as const).flatMap((orientation) =>
      ['in', 'out'].map((direction) => ({
        label:
          (orientation === 'horz' ? 'Horizontal' : 'Vertical') +
          (direction === 'in' ? ' In' : ' Out'),
        value: { effect, orientation, direction },
      })),
    );
  const choices: [string, string][] = ['push', 'wipe', 'cover', 'pull'].includes(effect)
    ? [
        ['l', 'From Right'],
        ['r', 'From Left'],
        ['u', 'From Bottom'],
        ['d', 'From Top'],
        ...(['cover', 'pull'].includes(effect)
          ? ([
              ['lu', 'From Bottom-Right'],
              ['ru', 'From Bottom-Left'],
              ['ld', 'From Top-Right'],
              ['rd', 'From Top-Left'],
            ] as [string, string][])
          : []),
      ]
    : ['blinds', 'checker', 'comb', 'randomBar'].includes(effect)
      ? [
          ['horz', 'Horizontal'],
          ['vert', 'Vertical'],
        ]
      : effect === 'zoom'
        ? [
            ['in', 'In'],
            ['out', 'Out'],
          ]
        : effect === 'strips'
          ? [
              ['lu', 'From Bottom-Right'],
              ['ru', 'From Bottom-Left'],
              ['ld', 'From Top-Right'],
              ['rd', 'From Top-Left'],
            ]
          : [];
  return choices.map(([direction, label]) => ({ label, value: { effect, direction } }));
}
export function transitionTile(effect: string, label: string) {
  const drawing =
    effect === 'none'
      ? '<path d="M8 7h32v22H8z" fill="none" stroke="currentColor"/>'
      : effect === 'fade'
        ? '<path d="M8 7h16v22H8z" fill="currentColor" opacity=".3"/><path d="M24 7h16v22H24z" fill="currentColor" opacity=".8"/>'
        : effect === 'split'
          ? '<path d="M8 7h14v22H8zM26 7h14v22H26z" fill="currentColor"/><path d="m19 18-5-4v8zm10 0 5-4v8z" fill="white"/>'
          : '<path d="M8 7h32v22H8z" fill="currentColor" opacity=".6"/><path d="M34 18H15m6-6-6 6 6 6" fill="none" stroke="white" stroke-width="2"/>';
  return `<button class="transition-tile" data-edit="transition-effect" data-effect="${effect}" aria-pressed="false"><svg viewBox="0 0 48 36" aria-hidden="true">${drawing}</svg><span>${label}</span></button>`;
}
