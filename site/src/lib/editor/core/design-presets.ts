// Office's built-in color sets and font pairs, as listed in PowerPoint's
// Design ▸ Colors and Design ▸ Fonts menus. A theme in the gallery is a color
// set with a font pair; the library rewrites the deck theme's `a:clrScheme`
// and `a:fontScheme` in place, keeping its masters and layouts.

import type { PresentationFontsInput, PresentationTheme } from '@office-kit/pptx';

export type ColorSet = Omit<PresentationTheme, 'dark1' | 'light1'>;

const set = (
  name: string,
  dark2: string,
  light2: string,
  accents: readonly string[],
  hyperlink: string,
  followedHyperlink: string,
): ColorSet => {
  const [accent1, accent2, accent3, accent4, accent5, accent6] = accents as [
    string,
    string,
    string,
    string,
    string,
    string,
  ];
  return {
    name,
    dark2,
    light2,
    accent1,
    accent2,
    accent3,
    accent4,
    accent5,
    accent6,
    hyperlink,
    followedHyperlink,
  };
};

export const COLOR_SETS: readonly ColorSet[] = [
  set(
    'Office',
    '#0E2841',
    '#E8E8E8',
    ['#156082', '#E97132', '#196B24', '#0F9ED5', '#A02B93', '#4EA72E'],
    '#467886',
    '#96607D',
  ),
  set(
    'Office 2013 - 2022',
    '#44546A',
    '#E7E6E6',
    ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'],
    '#0563C1',
    '#954F72',
  ),
  set(
    'Office 2007 - 2010',
    '#1F497D',
    '#EEECE1',
    ['#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646'],
    '#0000FF',
    '#800080',
  ),
  set(
    'Grayscale',
    '#000000',
    '#F8F8F8',
    ['#DDDDDD', '#B2B2B2', '#969696', '#808080', '#5F5F5F', '#4D4D4D'],
    '#5F5F5F',
    '#919191',
  ),
  set(
    'Blue',
    '#17406D',
    '#DBEFF9',
    ['#0F6FC6', '#009DD9', '#0BD0D9', '#10CF9B', '#7CCA62', '#A5C249'],
    '#F49100',
    '#85DFD0',
  ),
  set(
    'Green',
    '#455F51',
    '#E3DED1',
    ['#549E39', '#8AB833', '#C0CF3A', '#029676', '#4AB5C4', '#0989B1'],
    '#6B9F25',
    '#BA6906',
  ),
  set(
    'Red',
    '#323232',
    '#E5C243',
    ['#A5300F', '#D55816', '#E19825', '#B19C7D', '#7F5F52', '#B27D49'],
    '#6B9F25',
    '#B26B02',
  ),
  set(
    'Violet',
    '#373545',
    '#DCD8DC',
    ['#AD84C6', '#8784C7', '#5D739A', '#6997AF', '#84ACB6', '#6F8183'],
    '#69A020',
    '#8C8C8C',
  ),
];

export interface FontPair extends PresentationFontsInput {
  readonly name: string;
  readonly majorLatin: string;
  readonly minorLatin: string;
}

export const FONT_PAIRS: readonly FontPair[] = [
  { name: 'Office', majorLatin: 'Aptos Display', minorLatin: 'Aptos' },
  { name: 'Office 2013 - 2022', majorLatin: 'Calibri Light', minorLatin: 'Calibri' },
  { name: 'Office 2007 - 2010', majorLatin: 'Cambria', minorLatin: 'Calibri' },
  { name: 'Arial', majorLatin: 'Arial', minorLatin: 'Arial' },
  { name: 'Corbel', majorLatin: 'Corbel', minorLatin: 'Corbel' },
  { name: 'Georgia', majorLatin: 'Georgia', minorLatin: 'Georgia' },
  { name: 'Times New Roman-Arial', majorLatin: 'Times New Roman', minorLatin: 'Arial' },
  { name: 'Trebuchet MS', majorLatin: 'Trebuchet MS', minorLatin: 'Trebuchet MS' },
];

export interface ThemePreset {
  readonly name: string;
  readonly colors: ColorSet;
  readonly fonts: FontPair;
}

const pick = <T extends { name: string }>(list: readonly T[], name: string): T =>
  list.find((item) => item.name === name)!;

export const THEMES: readonly ThemePreset[] = [
  { name: 'Office Theme', colors: pick(COLOR_SETS, 'Office'), fonts: pick(FONT_PAIRS, 'Office') },
  {
    name: 'Office 2013 - 2022 Theme',
    colors: pick(COLOR_SETS, 'Office 2013 - 2022'),
    fonts: pick(FONT_PAIRS, 'Office 2013 - 2022'),
  },
  {
    name: 'Office 2007 - 2010 Theme',
    colors: pick(COLOR_SETS, 'Office 2007 - 2010'),
    fonts: pick(FONT_PAIRS, 'Office 2007 - 2010'),
  },
  { name: 'Blue', colors: pick(COLOR_SETS, 'Blue'), fonts: pick(FONT_PAIRS, 'Corbel') },
  { name: 'Green', colors: pick(COLOR_SETS, 'Green'), fonts: pick(FONT_PAIRS, 'Trebuchet MS') },
  { name: 'Red', colors: pick(COLOR_SETS, 'Red'), fonts: pick(FONT_PAIRS, 'Georgia') },
  {
    name: 'Violet',
    colors: pick(COLOR_SETS, 'Violet'),
    fonts: pick(FONT_PAIRS, 'Times New Roman-Arial'),
  },
];

export const accentsOf = (
  colors: Pick<ColorSet, 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6'>,
): string[] => [
  colors.accent1,
  colors.accent2,
  colors.accent3,
  colors.accent4,
  colors.accent5,
  colors.accent6,
];
