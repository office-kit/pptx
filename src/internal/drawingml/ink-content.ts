import {
  NS,
  attr,
  elem,
  getAttrValue,
  firstChildElement,
  qname,
  type XmlElement,
} from '../xml/index.ts';

/** Keep the AlternateContent wrapper as the shape identity for delete/copy/z-order. */
export const inkContentPart = (element: XmlElement): XmlElement | null => {
  if (element.name.namespaceURI === NS.pml && element.name.localName === 'contentPart')
    return firstChildElement(element, qname('p14', 'nvContentPartPr', NS.p14)) ? element : null;
  if (element.name.namespaceURI !== NS.mc || element.name.localName !== 'AlternateContent')
    return null;
  for (const child of element.children) {
    if (
      child.kind !== 'element' ||
      child.name.namespaceURI !== NS.mc ||
      child.name.localName !== 'Choice'
    )
      continue;
    const content = firstChildElement(child, qname('p', 'contentPart', NS.pml));
    if (content && firstChildElement(content, qname('p14', 'nvContentPartPr', NS.p14)))
      return content;
  }
  return null;
};

export const inkFallbackPicture = (element: XmlElement): XmlElement | null => {
  if (!inkContentPart(element)) return null;
  const fallback = firstChildElement(element, qname('mc', 'Fallback', NS.mc));
  return fallback && firstChildElement(fallback, qname('p', 'pic', NS.pml));
};

export const inkTransform = (element: XmlElement): XmlElement | null => {
  const content = inkContentPart(element);
  return content && firstChildElement(content, qname('p14', 'xfrm', NS.p14));
};

/** Keep names, accessibility text, visibility and actions consistent in both branches. */
export const syncInkFallbackMetadata = (element: XmlElement): void => {
  const content = inkContentPart(element);
  const nativeNv = content && firstChildElement(content, qname('p14', 'nvContentPartPr', NS.p14));
  const native = nativeNv && firstChildElement(nativeNv, qname('p14', 'cNvPr', NS.p14));
  const picture = inkFallbackPicture(element);
  const fallbackNv = picture && firstChildElement(picture, qname('p', 'nvPicPr', NS.pml));
  const fallback = fallbackNv && firstChildElement(fallbackNv, qname('p', 'cNvPr', NS.pml));
  if (!native || !fallback) return;
  const names = new Set(['name', 'descr', 'title', 'hidden']);
  const isShared = (attribute: XmlElement['attrs'][number]) =>
    attribute.name.namespaceURI === '' && names.has(attribute.name.localName);
  fallback.attrs = [
    ...fallback.attrs.filter((attribute) => !isShared(attribute)),
    ...structuredClone(native.attrs.filter(isShared)),
  ];
  // Both alternatives belong to the same slide, so hyperlink and sound
  // relationship IDs stay valid in the compatibility branch as well.
  const isAction = (child: XmlElement['children'][number]) =>
    child.kind === 'element' &&
    child.name.namespaceURI === NS.dml &&
    ['hlinkClick', 'hlinkHover'].includes(child.name.localName);
  fallback.children = [
    ...structuredClone(native.children.filter(isAction)),
    ...fallback.children.filter((child) => !isAction(child)),
  ];
};

/** Called after native geometry edits so older readers see the same placement. */
export const syncInkFallbackTransform = (element: XmlElement): void => {
  const native = inkTransform(element);
  const picture = inkFallbackPicture(element);
  const properties = picture && firstChildElement(picture, qname('p', 'spPr', NS.pml));
  if (!native || !properties) return;
  const copy = structuredClone(native);
  copy.name = qname('a', 'xfrm', NS.dml);
  const old = firstChildElement(properties, copy.name);
  if (old) properties.children.splice(properties.children.indexOf(old), 1, copy);
  else properties.children.unshift(copy);
};

export const inkAspectRatioLocked = (element: XmlElement): boolean => {
  const content = inkContentPart(element);
  const nv = content && firstChildElement(content, qname('p14', 'nvContentPartPr', NS.p14));
  const properties = nv && firstChildElement(nv, qname('p14', 'cNvContentPartPr', NS.p14));
  const locks = properties && firstChildElement(properties, qname('a14', 'cpLocks', NS.a14));
  const value = locks && getAttrValue(locks, qname('', 'noChangeAspect', ''));
  return value === '1' || value === 'true';
};

export const setInkAspectRatioLocked = (element: XmlElement, locked: boolean): void => {
  const content = inkContentPart(element);
  const nv = content && firstChildElement(content, qname('p14', 'nvContentPartPr', NS.p14));
  if (!nv) throw new Error('Ink has no non-visual properties.');
  let properties = firstChildElement(nv, qname('p14', 'cNvContentPartPr', NS.p14));
  if (!properties) {
    properties = elem(qname('p14', 'cNvContentPartPr', NS.p14));
    nv.children.splice(1, 0, properties);
  }
  const writeLock = (parent: XmlElement, prefix: string, name: string, namespace: string) => {
    let locks = firstChildElement(parent, qname(prefix, name, namespace));
    if (!locks) {
      locks = elem(qname(prefix, name, namespace), { prefixDecls: new Map([[prefix, namespace]]) });
      parent.children.unshift(locks);
    }
    locks.attrs = locks.attrs.filter(
      (a) => !(a.name.namespaceURI === '' && a.name.localName === 'noChangeAspect'),
    );
    locks.attrs.push(attr(qname('', 'noChangeAspect', ''), locked ? '1' : '0'));
  };
  writeLock(properties, 'a14', 'cpLocks', NS.a14);
  const picture = inkFallbackPicture(element);
  const fallbackNv = picture && firstChildElement(picture, qname('p', 'nvPicPr', NS.pml));
  if (fallbackNv) {
    let fallbackProperties = firstChildElement(fallbackNv, qname('p', 'cNvPicPr', NS.pml));
    if (!fallbackProperties) {
      fallbackProperties = elem(qname('p', 'cNvPicPr', NS.pml));
      fallbackNv.children.splice(1, 0, fallbackProperties);
    }
    writeLock(fallbackProperties, 'a', 'picLocks', NS.dml);
  }
};
