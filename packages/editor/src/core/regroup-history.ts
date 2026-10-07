import {
  getGroupChildren,
  getSlidePartName,
  getSlideShapes,
  getShapeId,
  type SlideData,
  type SlideShapeData,
} from '@office-kit/pptx';

interface FormerGroup {
  readonly slide: string;
  readonly ids: readonly number[];
}

/** Mac PowerPoint remembers dissolved groups only within the editing session. */
export class RegroupHistory {
  records: readonly FormerGroup[] = [];

  forget(slide: SlideData, shapes: readonly SlideShapeData[]): void {
    const part = getSlidePartName(slide);
    const ids = new Set(shapes.map(getShapeId));
    this.records = this.records.flatMap((record) => {
      if (record.slide !== part) return [record];
      const remaining = record.ids.filter((id) => !ids.has(id));
      return remaining.length >= 2 ? [{ ...record, ids: remaining }] : [];
    });
  }

  remember(slide: SlideData, shapes: readonly SlideShapeData[]): void {
    this.forget(slide, shapes);
    if (shapes.length >= 2)
      this.records = [
        ...this.records,
        { slide: getSlidePartName(slide), ids: shapes.map(getShapeId) },
      ];
  }

  members(slide: SlideData, selectedIds: readonly number[]): SlideShapeData[] {
    const part = getSlidePartName(slide);
    const shapes = getSlideShapes(slide);
    const parents = new Map(shapes.map((shape) => [getShapeId(shape), null as number | null]));
    for (const shape of shapes) {
      for (const child of getGroupChildren(shape))
        parents.set(getShapeId(child), getShapeId(shape));
    }
    const membership = new Map<number, ReadonlySet<number>>();
    for (const record of this.records) {
      if (record.slide !== part) continue;
      const siblings = new Map<number | null, Set<number>>();
      for (const id of record.ids) {
        const parent = parents.get(id);
        if (parent === undefined) continue;
        let ids = siblings.get(parent);
        if (!ids) siblings.set(parent, (ids = new Set()));
        ids.add(id);
      }
      for (const ids of siblings.values()) {
        if (ids.size >= 2) for (const id of ids) membership.set(id, ids);
      }
    }
    const first = selectedIds.find((id) => membership.has(id));
    if (first === undefined) return [];
    const ids = membership.get(first)!;
    return shapes.filter((shape) => ids.has(getShapeId(shape)));
  }

  /** Drop deleted members before IDs can be reused by a later editing action. */
  prune(slides: readonly SlideData[]): void {
    if (this.records.length === 0) return;
    const live = new Map(
      slides.map((slide) => [
        getSlidePartName(slide),
        new Set(getSlideShapes(slide).map(getShapeId)),
      ]),
    );
    this.records = this.records.flatMap((record) => {
      const ids = record.ids.filter((id) => live.get(record.slide)?.has(id));
      return ids.length >= 2 ? [{ ...record, ids }] : [];
    });
  }
}
