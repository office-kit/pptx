import {
  detectAudioFormat,
  emptyRels,
  nextRelId,
  partName,
  resolveTarget,
} from '../../internal/opc/index.ts';
import { REL_TYPES } from '../../internal/presentationml/index.ts';
import {
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  insertChildByRank,
  NS,
  qname,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  SHAPE_SLIDE,
  SLIDE_PART_NAME,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { commitAndRefresh } from './_helpers.ts';
import { findCNvPr } from './shape-click-action.ts';

export type ShapeActionTrigger = 'click' | 'hover';
export interface ShapeActionSound {
  readonly sound: { readonly name: string; readonly bytes: Uint8Array } | null;
  readonly stopPrevious: boolean;
}
const linkName = (trigger: ShapeActionTrigger) => {
  if (trigger !== 'click' && trigger !== 'hover') throw new Error('Invalid action trigger.');
  return qname('a', trigger === 'click' ? 'hlinkClick' : 'hlinkHover', NS.dml);
};
const soundName = qname('a', 'snd', NS.dml);
const endSoundName = qname('', 'endSnd', '');

/** Reads an action's embedded sound and stop-previous-sounds setting. */
export const getShapeActionSound = (
  shape: SlideShapeData,
  trigger: ShapeActionTrigger,
): ShapeActionSound => {
  const name = linkName(trigger);
  const metadata = findCNvPr(shape);
  const link = metadata && firstChildElement(metadata, name);
  const result: ShapeActionSound = {
    sound: null,
    stopPrevious: link ? ['1', 'true'].includes(getAttrValue(link, endSoundName) ?? '') : false,
  };
  const sound = link && firstChildElement(link, soundName);
  if (!sound) return result;
  const id = getAttrValue(sound, qname('r', 'embed', NS.officeDocRels));
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const rel = pkg
    .getRels(slide[SLIDE_PART_NAME])
    ?.items.find(
      (item) => item.id === id && item.type === REL_TYPES.audio && item.targetMode !== 'External',
    );
  if (!rel) return result;
  const target = resolveTarget(slide[SLIDE_PART_NAME], rel.target);
  const part = pkg.getPart(target);
  if (!part) return result;
  return {
    ...result,
    sound: { name: getAttrValue(sound, qname('', 'name', '')) ?? '', bytes: part.data.slice() },
  };
};

/** Sets an action sound without changing its destination or the other trigger.
 * New embedded sounds must use the WAV format supported by DrawingML a:snd.
 * Removed/replaced media parts are retained because other objects may use them.
 */
export const setShapeActionSound = (
  shape: SlideShapeData,
  trigger: ShapeActionTrigger,
  value: ShapeActionSound,
): void => {
  const name = linkName(trigger);
  const metadata = findCNvPr(shape);
  if (!metadata) throw new Error('Shape has no action metadata.');
  if (!value || typeof value.stopPrevious !== 'boolean' || value.sound === undefined)
    throw new Error('Invalid action sound settings.');
  const sound = value.sound;
  if (
    sound !== null &&
    (typeof sound.name !== 'string' ||
      !(sound.bytes instanceof Uint8Array) ||
      detectAudioFormat(sound.bytes) !== 'wav')
  )
    throw new Error('Action sound must be a WAV file with a name.');
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  let id: string | undefined;
  if (sound) {
    let media = pkg.parts.find(
      (part) =>
        part.contentType.toLowerCase().includes('wav') &&
        part.data.length === sound.bytes.length &&
        part.data.every((byte, index) => byte === sound.bytes[index]),
    );
    if (!media) {
      let ordinal = 1;
      while (pkg.getPart(partName(`/ppt/media/actionSound${ordinal}.wav`))) ordinal++;
      media = pkg.addPart(
        partName(`/ppt/media/actionSound${ordinal}.wav`),
        'audio/x-wav',
        sound.bytes.slice(),
      );
    }
    const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
    // Absolute OPC targets also support nonstandard slide/media part locations.
    const existing = rels.items.find(
      (rel) =>
        rel.type === REL_TYPES.audio &&
        rel.targetMode !== 'External' &&
        resolveTarget(slide[SLIDE_PART_NAME], rel.target) === media.name,
    );
    id = existing?.id ?? nextRelId(rels.items.map((rel) => rel.id));
    if (!existing) {
      rels.items.push({ id, type: REL_TYPES.audio, target: media.name, targetMode: 'Internal' });
      pkg.setRels(slide[SLIDE_PART_NAME], rels);
    }
  }
  let link = firstChildElement(metadata, name);
  if (!link) {
    if (!sound && !value.stopPrevious) return;
    link = elem(name, { attrs: [attr(qname('r', 'id', NS.officeDocRels), '')] });
    insertChildByRank(metadata, link, (child) =>
      child.name.namespaceURI === NS.dml && child.name.localName === 'hlinkClick'
        ? 0
        : child.name.namespaceURI === NS.dml && child.name.localName === 'hlinkHover'
          ? 1
          : 2,
    );
  }
  link.attrs = link.attrs.filter(
    ({ name }) => !(name.namespaceURI === '' && name.localName === 'endSnd'),
  );
  if (value.stopPrevious) link.attrs.push(attr(endSoundName, '1'));
  link.children = link.children.filter(
    (child) =>
      !(
        child.kind === 'element' &&
        child.name.namespaceURI === NS.dml &&
        child.name.localName === 'snd'
      ),
  );
  if (sound && id)
    insertChildByRank(
      link,
      elem(soundName, {
        attrs: [
          attr(qname('r', 'embed', NS.officeDocRels), id),
          attr(qname('', 'name', ''), sound.name),
        ],
      }),
      (child) => (child.name.namespaceURI === NS.dml && child.name.localName === 'snd' ? 0 : 1),
    );
  commitAndRefresh(shape);
};
