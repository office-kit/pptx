// What the master views draw: the deck's own slide, notes and handout master
// placeholders where the deck has them. Where it has none (a deck without a
// notes or handout master, or a placeholder without a transform), the views
// fall back to the default masters' geometry, which is what the reference desktop app
// creates — and what the library writes when the master is first edited.
import {
  getHandoutMasterPlaceholders,
  getNotesMasterPlaceholders,
  getNotesPageSize,
  getPresentationTheme,
  getSlideLayoutBackground,
  getSlideLayoutPlaceholders,
  getSlideMasterBackground,
  getSlideMasterLayouts,
  getSlideMasterPartNames,
  getSlideMasterPlaceholders,
  type HandoutSlidesPerPage,
  type PresentationData,
  type SlideLayoutData,
  type SlideLayoutPlaceholder,
} from '@office-kit/pptx';

/** A rectangle as fractions of the slide or page. */
export interface Area {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

// The default slide master (16:9, 12192000 × 6858000 EMU).
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
// The default notes and handout page, 7.5 × 10 in.
const PAGE_WIDTH_PT = 540;
const PAGE_HEIGHT_PT = 720;

// The default notes master (6858000 × 9144000 EMU).
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
  sldImg: pageArea(685800, 1143000, 5486400, 3086100),
  body: pageArea(685800, 4400550, 5486400, 3600450),
  ftr: pageArea(0, 8685213, 2971800, 458788),
  sldNum: pageArea(3884613, 8685213, 2971800, 458788),
} as const satisfies Record<string, Area>;

// The default handout master shares the notes master's header and footer
// areas. Its six-slides-per-page frames, measured from the reference desktop app (Mac, 16) on a
// 540 × 720 pt page: two columns 221 pt wide at x = 38.5 and 281, three rows
// 125 pt tall at y = 88.5, 297.5 and 506.
export const HANDOUT_AREAS = {
  hdr: NOTES_AREAS.hdr,
  dt: NOTES_AREAS.dt,
  ftr: NOTES_AREAS.ftr,
  sldNum: NOTES_AREAS.sldNum,
} as const satisfies Record<string, Area>;
const SIX_PER_PAGE: readonly Area[] = [88.5, 297.5, 506].flatMap((y) =>
  [38.5, 281].map((x) => ({
    x: x / PAGE_WIDTH_PT,
    y: y / PAGE_HEIGHT_PT,
    w: 221 / PAGE_WIDTH_PT,
    h: 125 / PAGE_HEIGHT_PT,
  })),
);

/**
 * Where the handout master draws its slide frames for `perPage` slides on a
 * page of `page` (in any unit) with slides of `slideAspect`. Six per page on
 * the default portrait page is measured from the reference desktop app (Mac); the other
 * layouts fill the same area (x 38.5–502 pt, y 88.5–631 pt of the page) on a
 * grid with the same 21.5 pt gutter, and three per page keeps the left
 * column, as the reference desktop app leaves the right for note lines. The outline layout
 * is one text area.
 */
export function handoutSlideFrames(
  perPage: HandoutSlidesPerPage,
  page: { width: number; height: number },
  slideAspect: number,
): readonly Area[] {
  const portrait = page.height >= page.width;
  if (perPage === 6 && portrait) return SIX_PER_PAGE;
  // The same margins as page fractions, so a landscape page keeps them.
  const left = 38.5 / PAGE_WIDTH_PT;
  const top = 88.5 / PAGE_HEIGHT_PT;
  const right = 502 / PAGE_WIDTH_PT;
  const bottom = 631 / PAGE_HEIGHT_PT;
  if (perPage === 'outline') return [{ x: left, y: top, w: right - left, h: bottom - top }];
  const grid: Record<Exclude<HandoutSlidesPerPage, 'outline'>, [number, number]> = portrait
    ? { 1: [1, 1], 2: [1, 2], 3: [2, 3], 4: [2, 2], 6: [2, 3], 9: [3, 3] }
    : { 1: [1, 1], 2: [2, 1], 3: [3, 1], 4: [2, 2], 6: [3, 2], 9: [3, 3] };
  const [cols, rows] = grid[perPage];
  const gutter = 21.5 / PAGE_WIDTH_PT;
  const ratio = (slideAspect * page.height) / page.width;
  const cellW = (right - left - gutter * (cols - 1)) / cols;
  const cellH = (bottom - top) / rows;
  const w = Math.min(cellW, cellH * 0.9 * ratio);
  const h = w / ratio;
  const frames: Area[] = [];
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < (perPage === 3 && portrait ? 1 : cols); col++) {
      frames.push({
        x: left + col * (cellW + gutter) + (cellW - w) / 2,
        y: top + row * cellH + (cellH - h) / 2,
        w,
        h,
      });
    }
  }
  return frames;
}

/** The notes and handout page in points, from the deck's `<p:notesSz>`. */
export function pageSizePt(pres: PresentationData): { width: number; height: number } {
  const size = getNotesPageSize(pres);
  return { width: size.width / 12700, height: size.height / 12700 };
}

export type PagePlaceholderRole = 'hdr' | 'dt' | 'sldImg' | 'body' | 'ftr' | 'sldNum';

/** A notes or handout master placeholder as fractions of the page. */
export interface PageBox extends Area {
  readonly role: PagePlaceholderRole;
}

const PAGE_ROLES = new Set<string>(['hdr', 'dt', 'sldImg', 'body', 'ftr', 'sldNum']);

/**
 * The notes or handout master's placeholders as page fractions: the deck's
 * own when it has the master, else the reference desktop app's default master.
 */
export function pageMasterBoxes(
  pres: PresentationData,
  kind: 'handoutMaster' | 'notesMaster',
): PageBox[] {
  const placeholders =
    kind === 'handoutMaster'
      ? getHandoutMasterPlaceholders(pres)
      : getNotesMasterPlaceholders(pres);
  const defaults = kind === 'handoutMaster' ? HANDOUT_AREAS : NOTES_AREAS;
  if (placeholders === null)
    return Object.entries(defaults).map(([role, area]) => ({
      role: role as PagePlaceholderRole,
      ...area,
    }));
  const page = getNotesPageSize(pres);
  return placeholders.flatMap((placeholder) => {
    const role = placeholder.type ?? '';
    if (!PAGE_ROLES.has(role)) return [];
    const bounds = placeholder.bounds;
    const fallback = (defaults as Record<string, Area>)[role];
    const area = bounds
      ? {
          x: bounds.x / page.width,
          y: bounds.y / page.height,
          w: bounds.w / page.width,
          h: bounds.h / page.height,
        }
      : fallback;
    return area ? [{ role: role as PagePlaceholderRole, ...area }] : [];
  });
}

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
    case 'body':
      return 'body';
    // `<p:ph>` without a type is an object placeholder (ECMA-376 §19.3.1.36).
    case null:
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

const toArea = (
  placeholder: SlideLayoutPlaceholder,
  size: { width: number; height: number },
): Area | null =>
  placeholder.bounds
    ? {
        x: placeholder.bounds.x / size.width,
        y: placeholder.bounds.y / size.height,
        w: placeholder.bounds.w / size.width,
        h: placeholder.bounds.h / size.height,
      }
    : null;

/**
 * The slide master's own placeholders (title, five-level body, the three
 * footers) as fractions of the slide, where the deck's master puts them.
 */
export function masterBoxes(
  pres: PresentationData,
  master: string,
  size: { width: number; height: number },
): MasterBox[] {
  return getSlideMasterPlaceholders(pres, master).flatMap((placeholder) => {
    const role = placeholderRole(placeholder.type);
    const area = toArea(placeholder, size) ?? (MASTER_AREAS as Record<string, Area>)[role];
    return area ? [{ role, type: placeholder.type, ...area }] : [];
  });
}

/**
 * A layout's placeholders as fractions of the slide; a placeholder without a
 * transform sits where its master's placeholder of the same kind does.
 */
export function layoutBoxes(
  layout: SlideLayoutData,
  size: { width: number; height: number },
  master: readonly MasterBox[],
): MasterBox[] {
  const inherited = (role: PlaceholderRole): Area =>
    master.find(
      (box) =>
        box.role === (role === 'subtitle' || role === 'object' || role === 'other' ? 'body' : role),
    ) ??
    (MASTER_AREAS as Record<string, Area>)[role] ??
    MASTER_AREAS.body;
  return getSlideLayoutPlaceholders(layout).map((placeholder) => {
    const role = placeholderRole(placeholder.type);
    const area = toArea(placeholder, size) ?? inherited(role);
    return { role, type: placeholder.type, x: area.x, y: area.y, w: area.w, h: area.h };
  });
}

export interface MasterGroup {
  readonly partName: string;
  readonly layouts: readonly SlideLayoutData[];
}

/** Masters in presentation order, each with its layouts in the master's order. */
export function masterGroups(pres: PresentationData): MasterGroup[] {
  return getSlideMasterPartNames(pres).map((partName) => ({
    partName,
    layouts: getSlideMasterLayouts(pres, partName),
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

/** Text that contrasts with `background`, as the reference desktop app's prompt text does. */
export function contrastInk(background: string, dark = '#000000', light = '#ffffff'): string {
  const hex = background.replace('#', '');
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  return 0.299 * r! + 0.587 * g! + 0.114 * b! > 140 ? dark : light;
}
