import { partName, relsPartNameFor, resolveTarget, type PartName } from '../opc/index.ts';
import { NS, parseXml, serializeXml } from '../xml/index.ts';
import type { OpcPackage, Part } from './package.ts';

/** Clone owned dependencies without interpreting their bodies or relationship types. */
export function duplicatePartGraph(
  pkg: OpcPackage,
  source: PartName,
  destination: PartName,
  sharedTypes: ReadonlySet<string>,
): void {
  const parts = new Map(pkg.parts.map((part) => [part.name.toLowerCase(), part]));
  const occupied = new Set(parts.keys());
  const copies = new Map<string, PartName>([[source.toLowerCase(), destination]]);
  const pending = [source];
  const planned: Part[] = [];
  const encoder = new TextEncoder();
  const decoder = new TextDecoder();
  occupied.add(destination.toLowerCase());
  const counters = new Map<string, number>();
  const allocate = (name: PartName): PartName => {
    const dot = name.lastIndexOf('.');
    const split = dot > name.lastIndexOf('/') ? dot : name.length;
    const stem = name.slice(0, split);
    const extension = name.slice(split);
    let n = counters.get(name.toLowerCase()) ?? 1;
    let candidate: PartName;
    do {
      candidate = partName(`${stem}-copy${n++}${extension}`);
    } while (occupied.has(candidate.toLowerCase()));
    counters.set(name.toLowerCase(), n);
    occupied.add(candidate.toLowerCase());
    return candidate;
  };

  for (let index = 0; index < pending.length; index++) {
    const name = pending[index]!;
    const part = parts.get(name.toLowerCase());
    if (!part) throw new Error(`Cannot duplicate missing dependency ${name}`);
    const copy = copies.get(name.toLowerCase())!;
    planned.push({ name: copy, contentType: part.contentType, data: new Uint8Array(part.data) });
    const relsPart = parts.get(relsPartNameFor(name).toLowerCase());
    if (!relsPart) continue;
    // Retain attributes and extension XML in relationships too. Only Target changes.
    const doc = parseXml(decoder.decode(relsPart.data));
    for (const element of doc.root.children) {
      if (
        element.kind !== 'element' ||
        element.name.localName !== 'Relationship' ||
        element.name.namespaceURI !== NS.relationships
      )
        continue;
      const target = element.attrs.find(
        (attribute) => attribute.name.namespaceURI === '' && attribute.name.localName === 'Target',
      );
      const type = element.attrs.find(
        (attribute) => attribute.name.namespaceURI === '' && attribute.name.localName === 'Type',
      )?.value;
      const mode = element.attrs.find(
        (attribute) =>
          attribute.name.namespaceURI === '' && attribute.name.localName === 'TargetMode',
      )?.value;
      if (!target || !type) throw new Error(`Invalid relationship in ${relsPart.name}`);
      if (mode === 'External') continue;
      const resolved = resolveTarget(name, target.value);
      if (!parts.has(resolved.toLowerCase()))
        throw new Error(`Cannot duplicate missing dependency ${resolved}`);
      let mapped = copies.get(resolved.toLowerCase());
      if (!mapped && !sharedTypes.has(type)) {
        mapped = allocate(resolved);
        copies.set(resolved.toLowerCase(), mapped);
        pending.push(resolved);
      }
      // Absolute part names stay correct even if the dependency has another directory.
      element.attrs = element.attrs.map((attribute) =>
        attribute === target ? { ...attribute, value: mapped ?? resolved } : attribute,
      );
    }
    planned.push({
      name: relsPartNameFor(copy),
      contentType: relsPart.contentType,
      data: encoder.encode(serializeXml(doc)),
    });
  }
  // Missing dependencies must fail before changing the original package.
  const names = new Set(parts.keys());
  for (const part of planned) {
    const key = part.name.toLowerCase();
    if (names.has(key)) throw new Error(`Cannot duplicate into existing part ${part.name}`);
    names.add(key);
  }
  for (const part of planned) pkg.addPart(part.name, part.contentType, part.data);
}
