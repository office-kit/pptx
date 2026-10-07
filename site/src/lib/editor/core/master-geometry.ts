// What the master views draw. The library exposes layout placeholders (with
// their own transforms) and master/layout backgrounds, but not the slide
// master's placeholders or the notes and handout masters. Where a part is not
// exposed, the master views fall back to the default Office masters' geometry,
// which is what PowerPoint creates for a new presentation.
import {
  getPresentationTheme,
  getSlideLayoutBackground,
  getSlideLayoutPartName,
  getSlideLayoutPlaceholders,
  getSlideLayouts,
  getSlideMasterBackground,
  getSlideMasterPartName,
  getSlideMasterPartNames,
  type PresentationData,
  type SlideLayoutData,
} from '@office-kit/pptx';

/** A rectangle as fractions of the slide or page. */
export interface Area {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

// The default Office slide master (16:9, 12192000 × 6858000 EMU).
export const MASTER_AREAS = {
  title: {
    x: 838200 / 12192000,
    y: 365125 / 6858000,
    w: 10515600 / 12192000,
    h: 1325563 / 6858000,
  },
  body: {
    x: 838200 / 12192000,
    y: 1825625 / 6858000,
    w: 10515600 / 12192000,
    h: 4351338 / 6858000,
  },
  dt: { x: 838200 / 12192000, y: 6356350 / 6858000, w: 2743200 / 12192000, h: 365125 / 6858000 },
  ftr: { x: 4038600 / 12192000, y: 6356350 / 6858000, w: 4114800 / 12192000, h: 365125 / 6858000 },
  sldNum: {
    x: 8610600 / 12192000,
    y: 6356350 / 6858000,
    w: 2743200 / 12192000,
    h: 365125 / 6858000,
  },
} as const satisfies Record<string, Area>;

/** Notes and handout pages are 7.5 × 10 in portrait unless the deck says otherwise. */
export const PAGE_WIDTH_PT = 540;
export const PAGE_HEIGHT_PT = 720;

// The default Office notes master (6858000 × 9144000 EMU).
const PAGE_W = 6858000;
const PAGE_H = 9144000;
const pageArea = (x: number, y: number, w: number, h: number): Area => ({
  x: x / PAGE_W,
  y: y / PAGE_H,
  w: w / PAGE_W,
  h: h / PAGE_H,
});
export const NOTES_AREAS = {
  hdr: pageArea(0, 0, 2971800, 458788),
  dt: pageArea(3884613, 0, 2971800, 458788),
  slideImage: pageArea(685800, 1143000, 5486400, 3086100),
  body: pageArea(685800, 4400550, 5486400, 3600450),
  ftr: pageArea(0, 8685213, 2971800, 458788),
  sldNum: pageArea(3884613, 8685213, 2971800, 458788),
} as const satisfies Record<string, Area>;

// The default handout master shares the notes master's header and footer
// areas. Its six-slides-per-page frames, measured from Mac PowerPoint 16 on a
// 540 × 720 pt page: two columns 221 pt wide at x = 38.5 and 281, three rows
// 125 pt tall at y = 88.5, 297.5 and 506.
export const HANDOUT_AREAS = {
  hdr: NOTES_AREAS.hdr,
  dt: NOTES_AREAS.dt,
  ftr: NOTES_AREAS.ftr,
  sldNum: NOTES_AREAS.sldNum,
} as const satisfies Record<string, Area>;
export const HANDOUT_SLIDE_FRAMES: readonly Area[] = [88.5, 297.5, 506].flatMap((y) =>
  [38.5, 281].map((x) => ({
    x: x / PAGE_WIDTH_PT,
    y: y / PAGE_HEIGHT_PT,
    w: 221 / PAGE_WIDTH_PT,
    h: 125 / PAGE_HEIGHT_PT,
  })),
);

export type PlaceholderRole =
  | 'title'
  | 'subtitle'
  | 'body'
  | 'object'
  | 'dt'
  | 'ftr'
  | 'sldNum'
  | 'other';

export function placeholderRole(type: string | null): PlaceholderRole {
  switch (type) {
    case 'title':
    case 'ctrTitle':
      return 'title';
    case 'subTitle':
      return 'subtitle';
    case 'dt':
    case 'ftr':
    case 'sldNum':
      return type;
    // `<p:ph>` without a type is a body placeholder (ECMA-376 §19.3.1.36).
    case null:
    case 'body':
      return 'body';
    case 'obj':
      return 'object';
    default:
      return 'other';
  }
}

export interface MasterBox extends Area {
  readonly role: PlaceholderRole;
  /** `<p:ph type>`: a title slide's `ctrTitle` is centred, unlike `title`. */
  readonly type: string | null;
}

const ROLE_AREA: Partial<Record<PlaceholderRole, Area>> = {
  title: MASTER_AREAS.title,
  subtitle: MASTER_AREAS.body,
  body: MASTER_AREAS.body,
  object: MASTER_AREAS.body,
  other: MASTER_AREAS.body,
  dt: MASTER_AREAS.dt,
  ftr: MASTER_AREAS.ftr,
  sldNum: MASTER_AREAS.sldNum,
};

/** The slide master's own placeholders: title, five-level body and the three footers. */
export const MASTER_BOXES: readonly MasterBox[] = (
  ['title', 'body', 'dt', 'ftr', 'sldNum'] as const
).map((role) => ({ role, type: role, ...MASTER_AREAS[role] }));

/** A layout's placeholders as fractions of the slide; untransformed ones sit where the master puts them. */
export function layoutBoxes(
  layout: SlideLayoutData,
  size: { width: number; height: number },
): MasterBox[] {
  return getSlideLayoutPlaceholders(layout).map((placeholder) => {
    const role = placeholderRole(placeholder.type);
    const bounds = placeholder.bounds;
    const type = placeholder.type;
    return bounds
      ? {
          role,
          type,
          x: bounds.x / size.width,
          y: bounds.y / size.height,
          w: bounds.w / size.width,
          h: bounds.h / size.height,
        }
      : { role, type, ...ROLE_AREA[role]! };
  });
}

export interface MasterGroup {
  readonly partName: string;
  readonly layouts: readonly SlideLayoutData[];
}

const partNumber = (name: string) => Number(/(\d+)\.xml$/.exec(name)?.[1] ?? 0);

/**
 * Masters in presentation order, each with its layouts. The library lists
 * layouts in package order and does not expose the master's layout list, so
 * layouts follow their part numbers, which PowerPoint assigns in master order.
 */
export function masterGroups(pres: PresentationData): MasterGroup[] {
  const layouts = [...getSlideLayouts(pres)].sort(
    (a, b) => partNumber(getSlideLayoutPartName(a)) - partNumber(getSlideLayoutPartName(b)),
  );
  return getSlideMasterPartNames(pres).map((partName) => ({
    partName,
    layouts: layouts.filter((layout) => getSlideMasterPartName(layout) === partName),
  }));
}

// Scheme colors as the default color map (bg1 = lt1, tx1 = dk1) resolves them.
const SCHEME_SLOTS = {
  bg1: 'light1',
  lt1: 'light1',
  tx1: 'dark1',
  dk1: 'dark1',
  bg2: 'light2',
  lt2: 'light2',
  tx2: 'dark2',
  dk2: 'dark2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
} as const;

/** A layout's (or, with `master`, its master's) solid background; other fills show white. */
export function solidBackground(
  pres: PresentationData,
  layout: SlideLayoutData,
  master = false,
): string {
  const own = master ? { kind: 'inherit' as const } : getSlideLayoutBackground(layout);
  const effective = own.kind === 'inherit' ? getSlideMasterBackground(pres, layout) : own;
  if (effective.kind !== 'solid') return '#ffffff';
  const token = effective.color.replace(/^scheme:/, '');
  const slot = SCHEME_SLOTS[token as keyof typeof SCHEME_SLOTS];
  const resolved = slot ? getPresentationTheme(pres)?.[slot] : effective.color;
  return resolved && /^#?[0-9a-f]{6}$/i.test(resolved)
    ? resolved.startsWith('#')
      ? resolved
      : `#${resolved}`
    : '#ffffff';
}

/** Text that contrasts with `background`, as PowerPoint's prompt text does. */
export function contrastInk(background: string, dark = '#000000', light = '#ffffff'): string {
  const hex = background.replace('#', '');
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  return 0.299 * r! + 0.587 * g! + 0.114 * b! > 140 ? dark : light;
}
