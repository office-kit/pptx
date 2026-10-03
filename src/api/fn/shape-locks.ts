import {
  NS,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import { SHAPE_ELEMENT, SHAPE_SLIDE, type SlideShapeData } from '../_internal-symbols.ts';
import { commitSlideData, refreshSlideData } from './_helpers.ts';

const GEOMETRY_LOCKS = ['noGrp', 'noRot', 'noMove', 'noResize'] as const;
const DRAWING_LOCKS = [
  ...GEOMETRY_LOCKS,
  'noEditPoints',
  'noAdjustHandles',
  'noChangeArrowheads',
  'noChangeShapeType',
];
const LOCKS = new Map([
  ['sp', { parent: 'nvSpPr', properties: 'cNvSpPr', element: 'spLocks', flags: DRAWING_LOCKS }],
  [
    'cxnSp',
    { parent: 'nvCxnSpPr', properties: 'cNvCxnSpPr', element: 'cxnSpLocks', flags: DRAWING_LOCKS },
  ],
  ['pic', { parent: 'nvPicPr', properties: 'cNvPicPr', element: 'picLocks', flags: DRAWING_LOCKS }],
  [
    'grpSp',
    {
      parent: 'nvGrpSpPr',
      properties: 'cNvGrpSpPr',
      element: 'grpSpLocks',
      flags: [...GEOMETRY_LOCKS, 'noUngrp'],
    },
  ],
  [
    'graphicFrame',
    {
      parent: 'nvGraphicFramePr',
      properties: 'cNvGraphicFramePr',
      element: 'graphicFrameLocks',
      flags: ['noGrp', 'noMove', 'noResize'],
    },
  ],
]);
const NO_CHANGE_ASPECT = qname('', 'noChangeAspect', '');

function lockProperties(shape: SlideShapeData) {
  const config = LOCKS.get(shape[SHAPE_ELEMENT].name.localName);
  if (!config) return null;
  const parent = firstChildElement(shape[SHAPE_ELEMENT], qname('p', config.parent, NS.pml));
  const properties = parent && firstChildElement(parent, qname('p', config.properties, NS.pml));
  return properties ? { config, properties } : null;
}

/** Whether the object's movement and resizing are both locked, as in PowerPoint's Selection Pane. */
export const isShapeLocked = (shape: SlideShapeData): boolean => {
  const target = lockProperties(shape);
  if (!target) return false;
  const locks = firstChildElement(target.properties, qname('a', target.config.element, NS.dml));
  return (
    !!locks &&
    ['noMove', 'noResize'].every((key) => {
      const value = getAttrValue(locks, qname('', key, ''));
      return value === '1' || value === 'true';
    })
  );
};

/** Whether PowerPoint preserves this object's aspect ratio while resizing it. */
export const isShapeAspectRatioLocked = (shape: SlideShapeData): boolean => {
  const target = lockProperties(shape);
  if (!target) return false;
  const locks = firstChildElement(target.properties, qname('a', target.config.element, NS.dml));
  if (!locks) return false;
  const value = getAttrValue(locks, NO_CHANGE_ASPECT);
  return value === '1' || value === 'true';
};

/**
 * Locks or unlocks object geometry like PowerPoint's Selection Pane. Text,
 * selection, aspect-ratio constraints and extension data remain unchanged.
 * Group children keep their own locks; pass all descendants for Lock All.
 * Arrays are committed once per slide, including selections across slides.
 */
export const setShapeLocked = (
  shapes: SlideShapeData | readonly SlideShapeData[],
  locked: boolean,
): void => {
  const targets = (Array.isArray(shapes) ? shapes : [shapes]).map((shape: SlideShapeData) => {
    const target = lockProperties(shape);
    if (!target)
      throw new Error('setShapeLocked: shape has no supported nonvisual drawing properties');
    return { shape, ...target };
  });
  for (const { config, properties } of targets) {
    let locks = firstChildElement(properties, qname('a', config.element, NS.dml));
    if (!locks) {
      if (!locked) continue;
      locks = elem(qname('a', config.element, NS.dml));
      properties.children.unshift(locks);
    }
    // Mac PowerPoint locks geometry but deliberately leaves text editable.
    const flags = new Set(config.flags);
    locks.attrs = locks.attrs.filter(
      (a) => a.name.namespaceURI !== '' || !flags.has(a.name.localName),
    );
    if (locked) for (const key of flags) locks.attrs.push(attr(qname('', key, ''), '1'));
    if (!locks.attrs.length && !locks.children.length)
      properties.children = properties.children.filter((child) => child !== locks);
  }
  for (const slide of new Set(targets.map(({ shape }) => shape[SHAPE_SLIDE]))) {
    commitSlideData(slide);
    refreshSlideData(slide);
  }
};

/**
 * Sets the DrawingML `noChangeAspect` constraint without disturbing other
 * nonvisual properties. An absent constraint and an explicit false value both
 * read as unlocked; disabling the constraint removes only this attribute.
 */
export const setShapeAspectRatioLocked = (
  shapes: SlideShapeData | readonly SlideShapeData[],
  locked: boolean,
): void => {
  const targets = (Array.isArray(shapes) ? shapes : [shapes]).map((shape: SlideShapeData) => {
    const target = lockProperties(shape);
    if (!target)
      throw new Error(
        'setShapeAspectRatioLocked: shape has no supported nonvisual drawing properties',
      );
    return { shape, ...target };
  });
  for (const { config, properties } of targets) {
    let locks = firstChildElement(properties, qname('a', config.element, NS.dml));
    if (!locks) {
      if (!locked) continue;
      locks = elem(qname('a', config.element, NS.dml));
      properties.children.unshift(locks);
    }
    locks.attrs = locks.attrs.filter(
      (value) => value.name.namespaceURI !== '' || value.name.localName !== 'noChangeAspect',
    );
    if (locked) locks.attrs.push(attr(NO_CHANGE_ASPECT, '1'));
    if (!locks.attrs.length && !locks.children.length)
      properties.children = properties.children.filter((child) => child !== locks);
  }
  for (const slide of new Set(targets.map(({ shape }) => shape[SHAPE_SLIDE]))) {
    commitSlideData(slide);
    refreshSlideData(slide);
  }
};
