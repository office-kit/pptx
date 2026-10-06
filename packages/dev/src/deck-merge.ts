import { strFromU8, strToU8 } from 'fflate';
// The core's package and XML layers are bundled into this package's build;
// they are not part of the `@office-kit/pptx` public API.
import {
  emptyRels,
  fromZipPath,
  parseContentTypes,
  parseRels,
  readZip,
  resolveTarget,
  serializeContentTypes,
  serializeRels,
  toZipPath,
  writeZip,
  type ContentTypes,
  type Relationship,
} from '../../../src/internal/opc/index.ts';
import {
  NS,
  elem,
  parseXml,
  serializeFragment,
  serializeXml,
  walkElements,
  type XmlAttr,
  type XmlElement,
} from '../../../src/internal/xml/index.ts';
import { comparablePart } from './fingerprint.ts';

export interface DeckConflict {
  /** ZIP part name, e.g. `ppt/slides/slide2.xml`. */
  part: string;
  /** 1-based slide number when the part belongs to a slide, its notes or their relationships. */
  slide: number | null;
  /** The colliding shape when the conflict is inside a shape tree. */
  shape: { id: string; name: string } | null;
  reason: 'both-changed' | 'deleted-and-changed' | 'both-added';
}

export type DeckMerge = { ok: true; bytes: Uint8Array } | { ok: false; conflicts: DeckConflict[] };

type Outcome<T> = { ok: true; value: T } | { ok: false; conflicts: DeckConflict[] };

const CONTENT_TYPES = '[Content_Types].xml';
const PRESENTATION = 'ppt/presentation.xml';
// Parts whose root is `cSld` + `spTree`, merged shape by shape.
const SHAPE_TREE_PART =
  /^ppt\/(?:slides|slideLayouts|slideMasters|notesSlides|notesMasters|handoutMasters)\/[^/]+\.xml$/;
const RELATIONSHIPS_PART = /(?:^|\/)_rels\/[^/]*\.rels$/;
const SLIDE_RELATIONSHIP = `${NS.officeDocRels}/slide`;
const NOTES_RELATIONSHIP = `${NS.officeDocRels}/notesSlide`;
// Children of `p:spTree` that describe the tree itself rather than a shape.
const SHAPE_TREE_PROPERTIES = new Set(['nvGrpSpPr', 'grpSpPr', 'extLst']);

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let index = 0; index < a.length; index++) if (a[index] !== b[index]) return false;
  return true;
}

interface Entry<T> {
  key: string;
  /** Comparable form of the item; equal values mean an unchanged item. */
  value: string;
  item: T;
}

type Ordering =
  /** Children may appear in any order, e.g. z-order in a shape tree. */
  | 'free'
  /** Children follow an `xsd:sequence`, so two independent insertions at one spot are ambiguous. */
  | 'schema'
  /** Order carries no meaning, so reorderings are ignored. */
  | 'none';

interface EntryMerge<T> {
  ordering: Ordering;
  /** Merges an item changed on both sides; `null` when the item is atomic. */
  refine?: (key: string, base: T, ours: T, theirs: T) => Outcome<T> | null;
  /** `sample` is absent when the conflict is about the items' order. */
  conflict: (key: string, reason: DeckConflict['reason'], sample?: T) => DeckConflict;
}

/**
 * Three-way merge of keyed items. An item changed on one side only takes that
 * side; an item changed differently on both sides is a conflict.
 */
function mergeEntries<T>(
  base: Entry<T>[],
  ours: Entry<T>[],
  theirs: Entry<T>[],
  options: EntryMerge<T>,
): Outcome<T[]> {
  const byKey = (entries: Entry<T>[]) => new Map(entries.map((entry) => [entry.key, entry]));
  const [b, o, t] = [byKey(base), byKey(ours), byKey(theirs)];
  const keys = new Set([...t.keys(), ...o.keys(), ...b.keys()]);
  const chosen = new Map<string, T>();
  const conflicts: DeckConflict[] = [];
  const same = (x: Entry<T> | undefined, y: Entry<T> | undefined) => x?.value === y?.value;
  for (const key of keys) {
    const [be, oe, te] = [b.get(key), o.get(key), t.get(key)];
    let pick: Entry<T> | undefined;
    if (same(oe, be)) pick = te;
    else if (same(te, be) || same(oe, te)) pick = oe;
    else {
      const refined = be && oe && te ? options.refine?.(key, be.item, oe.item, te.item) : null;
      if (refined?.ok) chosen.set(key, refined.value);
      else if (refined) conflicts.push(...refined.conflicts);
      else {
        const reason = !be ? 'both-added' : !oe || !te ? 'deleted-and-changed' : 'both-changed';
        conflicts.push(options.conflict(key, reason, (oe ?? te ?? be)!.item));
      }
      continue;
    }
    if (pick) chosen.set(key, pick.item);
  }
  if (conflicts.length) return { ok: false, conflicts };
  const order = mergeOrder(
    base.map((entry) => entry.key),
    ours.map((entry) => entry.key),
    theirs.map((entry) => entry.key),
    new Set(chosen.keys()),
    options.ordering,
  );
  if (!order) return { ok: false, conflicts: [options.conflict('', 'both-changed')] };
  return { ok: true, value: order.map((key) => chosen.get(key)!) };
}

/** True when the keys both lists share appear in the same order in each. */
function sameRelativeOrder(a: string[], b: string[]): boolean {
  const [inA, inB] = [new Set(a), new Set(b)];
  const sharedA = a.filter((key) => inB.has(key));
  const sharedB = b.filter((key) => inA.has(key));
  return sharedA.every((key, index) => key === sharedB[index]);
}

/**
 * Orders the kept keys. The side that reordered (or the source when neither
 * did) gives the order; the other side's additions follow their nearest
 * preceding sibling. `null` means the order cannot be decided.
 */
function mergeOrder(
  base: string[],
  ours: string[],
  theirs: string[],
  keep: Set<string>,
  ordering: Ordering,
): string[] | null {
  const oursMoved = ordering !== 'none' && !sameRelativeOrder(base, ours);
  const theirsMoved = ordering !== 'none' && !sameRelativeOrder(base, theirs);
  if (oursMoved && theirsMoved && !sameRelativeOrder(ours, theirs)) return null;
  const [primary, secondary] = oursMoved ? [ours, theirs] : [theirs, ours];
  const kept = primary.filter((key) => keep.has(key));
  const placed = new Set(kept);
  const inSecondary = new Set(secondary);
  const next = new Map<string | null, string | undefined>(
    kept.map((key, index) => [index === 0 ? null : kept[index - 1]!, key]),
  );
  if (kept.length) next.set(kept.at(-1)!, undefined);
  // Each anchor gets at most one insertion: the next one chains to the last.
  const after = new Map<string | null, string>();
  let anchor: string | null = null;
  let chained = false;
  for (const key of secondary) {
    if (!keep.has(key)) continue;
    if (placed.has(key)) {
      anchor = key;
      chained = false;
      continue;
    }
    if (!chained) {
      const following = next.get(anchor);
      // Both sides inserted different elements at the same spot.
      if (ordering === 'schema' && following !== undefined && !inSecondary.has(following))
        return null;
    }
    after.set(anchor, key);
    anchor = key;
    chained = true;
  }
  const result: string[] = [];
  const emit = (key: string | null) => {
    for (let at: string | null | undefined = key; at !== undefined; at = after.get(at))
      if (at !== null) result.push(at);
  };
  emit(null);
  for (const key of kept) emit(key);
  return result;
}

interface PartContext {
  part: string;
}

const partConflict = (
  context: PartContext,
  reason: DeckConflict['reason'],
  shape: DeckConflict['shape'] = null,
): DeckConflict => ({ part: context.part, slide: null, shape, reason });

function shapeIdentity(element: XmlElement): { id: string; name: string } | null {
  let found: XmlElement | undefined;
  walkElements(element, (child) => {
    if (found) return false;
    if (child.name.localName === 'cNvPr') found = child;
    return undefined;
  });
  if (!found) return null;
  const value = (name: string) => found!.attrs.find((item) => item.name.localName === name)?.value;
  const id = value('id');
  return id === undefined ? null : { id, name: value('name') ?? '' };
}

type ContainerKind = 'root' | 'cSld' | 'spTree';

const elementKey = (element: XmlElement) =>
  `${element.name.namespaceURI} ${element.name.localName}`;

/** Keys an element's children, or `null` when they cannot be told apart. */
function childEntries(element: XmlElement, kind: ContainerKind): Entry<XmlElement>[] | null {
  const entries: Entry<XmlElement>[] = [];
  const keys = new Set<string>();
  const occurrences = new Map<string, number>();
  for (const child of element.children) {
    if (child.kind !== 'element') {
      // Whitespace between elements carries no meaning in these parts.
      if (child.kind === 'text' && !child.data.trim()) continue;
      return null;
    }
    let key: string;
    if (
      kind === 'spTree' &&
      !(child.name.namespaceURI === NS.pml && SHAPE_TREE_PROPERTIES.has(child.name.localName))
    ) {
      const identity = shapeIdentity(child);
      if (!identity) return null;
      key = `shape ${identity.id}`;
    } else {
      const name = elementKey(child);
      const occurrence = occurrences.get(name) ?? 0;
      occurrences.set(name, occurrence + 1);
      key = `${name} ${occurrence}`;
    }
    if (keys.has(key)) return null;
    keys.add(key);
    entries.push({ key, value: serializeFragment(child), item: child });
  }
  return entries;
}

const attributesValue = (attrs: XmlAttr[]) =>
  JSON.stringify(attrs.map((item) => [item.name.namespaceURI, item.name.localName, item.value]));

const NESTED: Partial<Record<ContainerKind, { localName: string; kind: ContainerKind }>> = {
  root: { localName: 'cSld', kind: 'cSld' },
  cSld: { localName: 'spTree', kind: 'spTree' },
};

/** Merges a slide-like element: its attributes as one unit, its children by key. */
function mergeContainer(
  context: PartContext,
  base: XmlElement,
  ours: XmlElement,
  theirs: XmlElement,
  kind: ContainerKind,
): Outcome<XmlElement> {
  const conflicts: DeckConflict[] = [];
  const [ba, oa, ta] = [base.attrs, ours.attrs, theirs.attrs].map(attributesValue);
  let attrs = theirs.attrs;
  if (oa !== ba && (ta === ba || oa === ta)) attrs = ours.attrs;
  else if (oa !== ba) conflicts.push(partConflict(context, 'both-changed'));
  // Shapes taken from the editor's copy may use prefixes declared on its root.
  const prefixDecls = new Map(theirs.prefixDecls);
  for (const [prefix, uri] of ours.prefixDecls) {
    const existing = prefixDecls.get(prefix);
    if (existing === undefined) prefixDecls.set(prefix, uri);
    else if (existing !== uri) conflicts.push(partConflict(context, 'both-changed'));
  }
  const entries = [base, ours, theirs].map((element) => childEntries(element, kind));
  if (entries.some((list) => !list)) conflicts.push(partConflict(context, 'both-changed'));
  if (conflicts.length) return { ok: false, conflicts };
  const nested = NESTED[kind];
  const children = mergeEntries(entries[0]!, entries[1]!, entries[2]!, {
    ordering: kind === 'spTree' ? 'free' : 'schema',
    refine: (_key, b, o, t) =>
      nested && t.name.namespaceURI === NS.pml && t.name.localName === nested.localName
        ? mergeContainer(context, b, o, t, nested.kind)
        : null,
    conflict: (key, reason, sample) =>
      partConflict(
        context,
        reason,
        sample && key.startsWith('shape ') ? shapeIdentity(sample) : null,
      ),
  });
  if (!children.ok) return children;
  return {
    ok: true,
    value: elem(theirs.name, { attrs, prefixDecls, children: children.value }),
  };
}

function mergeShapeTreePart(
  context: PartContext,
  base: Uint8Array,
  ours: Uint8Array,
  theirs: Uint8Array,
): Outcome<Uint8Array> {
  const [b, o, t] = [base, ours, theirs].map((bytes) => parseXml(strFromU8(bytes)));
  const root = mergeContainer(context, b!.root, o!.root, t!.root, 'root');
  if (!root.ok) return root;
  return { ok: true, value: strToU8(serializeXml({ ...t!, root: root.value })) };
}

function relationshipEntries(bytes: Uint8Array | undefined): Entry<Relationship>[] {
  const rels = bytes ? parseRels(strFromU8(bytes)) : emptyRels();
  return rels.items.map((item) => ({
    key: item.id,
    value: JSON.stringify([item.type, item.target, item.targetMode]),
    item,
  }));
}

function mergeRelationshipsPart(
  context: PartContext,
  base: Uint8Array | undefined,
  ours: Uint8Array,
  theirs: Uint8Array,
): Outcome<Uint8Array> {
  const items = mergeEntries(
    relationshipEntries(base),
    relationshipEntries(ours),
    relationshipEntries(theirs),
    { ordering: 'none', conflict: (_key, reason) => partConflict(context, reason) },
  );
  if (!items.ok) return items;
  return { ok: true, value: strToU8(serializeRels({ items: items.value })) };
}

/**
 * Content types are declarations about parts, so both sides' entries are kept
 * and only overrides for parts the merged package no longer has are dropped.
 */
function mergeContentTypes(
  ours: Uint8Array,
  theirs: Uint8Array,
  parts: Set<string>,
): Outcome<Uint8Array> {
  const [o, t] = [ours, theirs].map((bytes) => parseContentTypes(strFromU8(bytes)));
  const merged: ContentTypes = { defaults: [], overrides: [] };
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  const conflicts: DeckConflict[] = [];
  const context = { part: CONTENT_TYPES };
  for (const types of [t!, o!]) {
    for (const item of types.defaults) {
      const existing = defaults.get(item.extension);
      if (existing === undefined) {
        defaults.set(item.extension, item.contentType);
        merged.defaults.push(item);
      } else if (existing !== item.contentType)
        conflicts.push(partConflict(context, 'both-changed'));
    }
    for (const item of types.overrides) {
      const key = item.partName.toLowerCase();
      if (!parts.has(key)) continue;
      const existing = overrides.get(key);
      if (existing === undefined) {
        overrides.set(key, item.contentType);
        merged.overrides.push(item);
      } else if (existing !== item.contentType)
        conflicts.push(partConflict(context, 'both-changed'));
    }
  }
  if (conflicts.length) return { ok: false, conflicts };
  return { ok: true, value: strToU8(serializeContentTypes(merged)) };
}

/** Maps each slide's parts (the slide, its notes and their relationships) to its number. */
function slideNumbers(parts: Map<string, Uint8Array>): Map<string, number> {
  const numbers = new Map<string, number>();
  const presentation = parts.get(PRESENTATION);
  const presentationRels = parts.get('ppt/_rels/presentation.xml.rels');
  if (!presentation || !presentationRels) return numbers;
  const targets = new Map(
    parseRels(strFromU8(presentationRels))
      .items.filter((item) => item.type === SLIDE_RELATIONSHIP)
      .map((item) => [item.id, toZipPath(resolveTarget(fromZipPath(PRESENTATION), item.target))]),
  );
  let number = 0;
  walkElements(parseXml(strFromU8(presentation)).root, (element) => {
    if (element.name.localName !== 'sldId') return undefined;
    number++;
    const id = element.attrs.find(
      (item) => item.name.namespaceURI === NS.officeDocRels && item.name.localName === 'id',
    )?.value;
    const slide = id === undefined ? undefined : targets.get(id);
    if (!slide) return false;
    const relsPart = relsName(slide);
    numbers.set(slide, number).set(relsPart, number);
    const rels = parts.get(relsPart);
    if (!rels) return false;
    for (const item of parseRels(strFromU8(rels)).items) {
      if (item.type !== NOTES_RELATIONSHIP) continue;
      const notes = toZipPath(resolveTarget(fromZipPath(slide), item.target));
      numbers.set(notes, number).set(relsName(notes), number);
    }
    return false;
  });
  return numbers;
}

const relsName = (part: string) => {
  const slash = part.lastIndexOf('/');
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
};

/**
 * Both sides allocate the next free relationship ID, so a relationship the
 * editor added and a different one the source added often share an ID. IDs
 * are private to a part and its relationships, so the editor's is renamed
 * (in its relationships and in the part referring to it) instead of
 * reporting a conflict.
 */
function renumberAddedRelationships(
  base: Map<string, Uint8Array>,
  ours: Map<string, Uint8Array>,
  theirs: Map<string, Uint8Array>,
) {
  for (const [name, bytes] of ours) {
    const owner = relationshipsOwner(name);
    const theirsBytes = theirs.get(name);
    const ownerBytes = owner && ours.get(owner);
    if (!ownerBytes || !theirsBytes) continue;
    const parse = (value: Uint8Array | undefined) =>
      value ? parseRels(strFromU8(value)) : emptyRels();
    const [b, o, t] = [parse(base.get(name)), parse(bytes), parse(theirsBytes)];
    const inBase = new Set(b.items.map((item) => item.id));
    const theirsById = new Map(t.items.map((item) => [item.id, item]));
    const used = new Set([...inBase, ...theirsById.keys(), ...o.items.map((item) => item.id)]);
    const renamed = new Map<string, string>();
    let next = 1;
    for (const item of o.items) {
      const other = theirsById.get(item.id);
      if (inBase.has(item.id) || !other || sameRelationship(item, other)) continue;
      while (used.has(`rId${next}`)) next++;
      renamed.set(item.id, `rId${next}`);
      used.add(`rId${next}`);
      item.id = `rId${next}`;
    }
    if (!renamed.size) continue;
    ours.set(name, strToU8(serializeRels(o)));
    const document = parseXml(strFromU8(ownerBytes));
    walkElements(document.root, (element) => {
      element.attrs = element.attrs.map((item) =>
        item.name.namespaceURI === NS.officeDocRels && renamed.has(item.value)
          ? { name: item.name, value: renamed.get(item.value)! }
          : item,
      );
      return undefined;
    });
    ours.set(owner, strToU8(serializeXml(document)));
  }
}

const sameRelationship = (a: Relationship, b: Relationship) =>
  a.type === b.type && a.target === b.target && a.targetMode === b.targetMode;

/** The part whose relationships `name` holds; `null` for the package's own. */
function relationshipsOwner(name: string): string | null {
  const match = /^(.*?)_rels\/([^/]+)\.rels$/.exec(name);
  return match ? `${match[1]}${match[2]}` : null;
}

function unzip(bytes: Uint8Array): Map<string, Uint8Array> {
  return new Map(readZip(bytes).entries.map((entry) => [entry.name, entry.data]));
}

/**
 * Three-way merge of presentation packages: `base` is the source build the
 * editor's copy was made from, `ours` the editor's copy and `theirs` the new
 * source build. Changes on different parts, shapes or relationships combine;
 * the same item changed differently on both sides is reported as a conflict.
 */
export function mergeDecks(base: Uint8Array, ours: Uint8Array, theirs: Uint8Array): DeckMerge {
  const [b, o, t] = [unzip(base), unzip(ours), unzip(theirs)] as const;
  renumberAddedRelationships(b, o, t);
  const names = new Set([...t.keys(), ...o.keys(), ...b.keys()]);
  const result = new Map<string, Uint8Array>();
  const conflicts: DeckConflict[] = [];
  const same = (name: string, x: Uint8Array | undefined, y: Uint8Array | undefined) =>
    x === undefined || y === undefined
      ? x === y
      : sameBytes(comparablePart(name, x), comparablePart(name, y));
  for (const name of names) {
    if (name === CONTENT_TYPES) continue;
    const [bp, op, tp] = [b.get(name), o.get(name), t.get(name)];
    let pick: Uint8Array | undefined;
    if (same(name, op, bp)) pick = tp;
    else if (same(name, tp, bp) || same(name, op, tp)) pick = op;
    else {
      const context = { part: name };
      let merged: Outcome<Uint8Array> | null = null;
      if (op && tp && RELATIONSHIPS_PART.test(name))
        merged = mergeRelationshipsPart(context, bp, op, tp);
      else if (bp && op && tp && SHAPE_TREE_PART.test(name))
        merged = mergeShapeTreePart(context, bp, op, tp);
      if (merged?.ok) result.set(name, merged.value);
      else if (merged) conflicts.push(...merged.conflicts);
      else
        conflicts.push(
          partConflict(
            context,
            !bp ? 'both-added' : !op || !tp ? 'deleted-and-changed' : 'both-changed',
          ),
        );
      continue;
    }
    if (pick) result.set(name, pick);
  }
  const [bt, ot, tt] = [b, o, t].map((parts) => parts.get(CONTENT_TYPES)!);
  let contentTypes: Uint8Array | undefined;
  if (same(CONTENT_TYPES, ot, bt)) contentTypes = tt;
  else if (same(CONTENT_TYPES, tt, bt) || same(CONTENT_TYPES, ot, tt)) contentTypes = ot;
  else {
    const merged = mergeContentTypes(
      ot!,
      tt!,
      new Set([...result.keys()].map((name) => `/${name}`.toLowerCase())),
    );
    if (merged.ok) contentTypes = merged.value;
    else conflicts.push(...merged.conflicts);
  }
  if (conflicts.length) {
    const numbers = [b, t, o].reduce(
      (all, parts) => new Map([...all, ...slideNumbers(parts)]),
      new Map<string, number>(),
    );
    return {
      ok: false,
      conflicts: conflicts.map((conflict) => ({
        ...conflict,
        slide: numbers.get(conflict.part) ?? null,
      })),
    };
  }
  return {
    ok: true,
    bytes: writeZip([
      { name: CONTENT_TYPES, data: contentTypes! },
      ...[...result].map(([name, data]) => ({ name, data })),
    ]),
  };
}

/** One line per conflict, for the CLI and server messages. */
export function describeConflict(conflict: DeckConflict): string {
  const where = conflict.slide === null ? conflict.part : `Slide ${conflict.slide}`;
  const what = conflict.shape
    ? `${where}: ${conflict.shape.name || `shape ${conflict.shape.id}`}`
    : where;
  return conflict.reason === 'deleted-and-changed'
    ? `${what} was deleted on one side and changed on the other.`
    : `${what} was changed both in the editor and in the source.`;
}
