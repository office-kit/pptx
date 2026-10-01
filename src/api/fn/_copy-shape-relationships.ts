import { emptyRels, nextRelId, resolveTarget, type PartName } from '../../internal/opc/index.ts';
import { copyPartGraphs } from '../../internal/parts/duplicate-graph.ts';
import { NS, type XmlElement } from '../../internal/xml/index.ts';
import { INTERNAL_PACKAGE, SLIDE_PART_NAME, type SlideData } from '../_internal-symbols.ts';

/** Remap only relationships referenced by the copied XML, preserving shared parts. */
export function copyShapeRelationships(
  sourceSlide: SlideData,
  targetSlide: SlideData,
  cloned: XmlElement,
  caller: string,
): void {
  prepareShapeRelationshipCopy(sourceSlide)(targetSlide, cloned, caller);
}

/** Snapshot source relationships once when distributing text to several slides. */
export function prepareShapeRelationshipCopy(sourceSlide: SlideData) {
  const sourcePkg = sourceSlide[INTERNAL_PACKAGE];
  const sourceRels = sourcePkg.getRels(sourceSlide[SLIDE_PART_NAME]);
  const relsById = new Map(sourceRels?.items.map((rel) => [rel.id, rel]));
  return (targetSlide: SlideData, cloned: XmlElement, caller: string): void => {
    const pkg = targetSlide[INTERNAL_PACKAGE];
    const referencedIds = new Set<string>();
    rewriteRIdReferences(cloned, (id) => {
      referencedIds.add(id);
      return id;
    });
    const roots = new Map<PartName, null>();
    for (const id of referencedIds) {
      const rel = relsById.get(id);
      if (!rel) throw new Error(`${caller}: missing relationship ${id}`);
      if (rel.targetMode !== 'External')
        roots.set(resolveTarget(sourceSlide[SLIDE_PART_NAME], rel.target), null);
    }
    const copies =
      sourcePkg === pkg
        ? null
        : copyPartGraphs(
            sourcePkg,
            pkg,
            roots,
            new Set(),
            new Map([[sourceSlide[SLIDE_PART_NAME], targetSlide[SLIDE_PART_NAME]]]),
          );
    const targetRels = pkg.getRels(targetSlide[SLIDE_PART_NAME]) ?? emptyRels();
    const relIds = new Map<string, string>();
    const usedIds = targetRels.items.map((rel) => rel.id);
    const relationshipKey = (type: string, target: string, mode: string): string =>
      JSON.stringify([type, target, mode]);
    const existingRels = new Map(
      targetRels.items.map((rel) => [
        relationshipKey(
          rel.type,
          rel.targetMode === 'External'
            ? rel.target
            : resolveTarget(targetSlide[SLIDE_PART_NAME], rel.target),
          rel.targetMode,
        ),
        rel.id,
      ]),
    );
    const availableIds = nextRelId(usedIds, referencedIds.size);
    let allocated = 0;
    for (const id of referencedIds) {
      const rel = relsById.get(id)!;
      const resolved =
        rel.targetMode === 'External'
          ? rel.target
          : resolveTarget(sourceSlide[SLIDE_PART_NAME], rel.target);
      const target = copies?.get(resolved.toLowerCase()) ?? resolved;
      const key = relationshipKey(rel.type, target, rel.targetMode);
      const existing = existingRels.get(key);
      const newRId = existing ?? availableIds[allocated++]!;
      if (!existing) {
        targetRels.items.push({ ...rel, target, id: newRId });
        existingRels.set(key, newRId);
      }
      relIds.set(id, newRId);
    }
    rewriteRIdReferences(cloned, (id) => relIds.get(id)!);
    pkg.setRels(targetSlide[SLIDE_PART_NAME], targetRels);
  };
}

const rewriteRIdReferences = (root: XmlElement, map: (oldRId: string) => string): void => {
  const walk = (el: XmlElement): void => {
    el.attrs = el.attrs.map((a) => {
      if (
        a.value !== '' &&
        a.name.namespaceURI === NS.officeDocRels &&
        (a.name.localName === 'id' || a.name.localName === 'embed' || a.name.localName === 'link')
      ) {
        return { name: a.name, value: map(a.value) };
      }
      return a;
    });
    for (const c of el.children) {
      if (c.kind === 'element') walk(c);
    }
  };
  walk(root);
};
