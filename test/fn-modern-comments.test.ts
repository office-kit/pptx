// Modern comments — the `<p188:cm>` threads PowerPoint 2021 and Microsoft 365
// write ([MS-PPTX] §2.16.1, schema §5.14).
//
// There is no fixture from a real PowerPoint here, so the parts are written
// out the way that specification says PowerPoint writes them — including the
// things this library does not model (a shape anchor, an extension list, a
// reaction) so that leaving them alone can be checked rather than assumed.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addSlideComment,
  clearSlideComments,
  getCommentAuthor,
  getCommentDate,
  getCommentFormat,
  getCommentParent,
  getCommentStatus,
  getCommentText,
  getSlideComments,
  getSlides,
  loadPresentation,
  removeSlideComment,
  savePresentation,
  setCommentStatus,
  setCommentText,
  type PresentationData,
} from '../src/api/index.ts';
import { partName } from '../src/internal/opc/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const encode = (s: string): Uint8Array => new TextEncoder().encode(s);
const decode = (b: Uint8Array): string => new TextDecoder().decode(b);

const AUTHORS = '/ppt/authors.xml';
const COMMENTS = '/ppt/comments/modernComment1.xml';
const AUTHORS_TYPE = 'application/vnd.ms-powerpoint.authors+xml';
const COMMENTS_TYPE = 'application/vnd.ms-powerpoint.comments+xml';
const AUTHORS_REL = 'http://schemas.microsoft.com/office/2018/10/relationships/authors';
const COMMENTS_REL = 'http://schemas.microsoft.com/office/2018/10/relationships/comments';

const REVIEWER = '{A0000000-0000-0000-0000-000000000001}';
const SECOND = '{A0000000-0000-0000-0000-000000000002}';
const FIRST_THREAD = '{1D1EC5A0-0000-0000-0000-000000000001}';
const ONLY_REPLY = '{1D1EC5A0-0000-0000-0000-000000000002}';
const SECOND_THREAD = '{1D1EC5A0-0000-0000-0000-000000000003}';

const AUTHORS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:author id="${REVIEWER}" name="Reviewer" initials="R" userId="S::reviewer@example.com::0" providerId="AD"/><p188:author id="${SECOND}" name="校閲 二郎" initials="校" userId="S::jiro@example.com::0" providerId="AD"/></p188:authorLst>`;

// A resolved thread with one reply and an extension list we do not model,
// then an active thread anchored to a shape rather than to the slide.
const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command" xmlns:ac="http://schemas.microsoft.com/office/drawing/2013/main/command" xmlns:p223="http://schemas.microsoft.com/office/powerpoint/2022/03/main"><p188:cm id="${FIRST_THREAD}" authorId="${REVIEWER}" created="2026-05-15T12:00:00.000" status="resolved"><pc:sldMkLst><pc:docMk/><pc:sldMk cId="3141592653" sldId="256"/></pc:sldMkLst><p188:pos x="1270000" y="635000"/><p188:replyLst><p188:reply id="${ONLY_REPLY}" authorId="${SECOND}" created="2026-05-15T12:05:00.000" status="active"><p188:txBody><a:bodyPr/><a:p><a:r><a:t>承知しました。</a:t></a:r></a:p></p188:txBody></p188:reply></p188:replyLst><p188:txBody><a:bodyPr/><a:p><a:r><a:t>Tighten this headline.</a:t></a:r></a:p></p188:txBody><p188:extLst><p:ext uri="{9E7C0F2E-0000-0000-0000-00000000FEED}"><p223:reactions xmlns:p223="http://schemas.microsoft.com/office/powerpoint/2022/03/main"><p223:reaction authorId="${SECOND}" type="like"/></p223:reactions></p:ext></p188:extLst></p188:cm><p188:cm id="${SECOND_THREAD}" authorId="${SECOND}" created="2026-05-16T09:30:00.000"><ac:deMkLst><pc:docMk/><pc:sldMk cId="3141592653" sldId="256"/><ac:deMk id="7" idx="0"/></ac:deMkLst><p188:txBody><a:bodyPr/><a:p><a:r><a:t>This chart needs a source line.</a:t></a:r></a:p></p188:txBody></p188:cm></p188:cmLst>`;

/** A deck whose first slide carries the two threads above. */
const withModernComments = async (): Promise<PresentationData> => {
  const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
  const pkg = _internalPackageOf(pres);
  pkg.addPart(partName(AUTHORS), AUTHORS_TYPE, encode(AUTHORS_XML));
  pkg.addPart(partName(COMMENTS), COMMENTS_TYPE, encode(COMMENTS_XML));
  const presRels = pkg.getRels(partName('/ppt/presentation.xml'));
  if (!presRels) throw new Error('expected presentation rels');
  presRels.items.push({
    id: 'rId900',
    type: AUTHORS_REL,
    target: AUTHORS,
    targetMode: 'Internal',
  });
  pkg.setRels(partName('/ppt/presentation.xml'), presRels);
  const slideRels = pkg.getRels(partName('/ppt/slides/slide1.xml'));
  if (!slideRels) throw new Error('expected slide rels');
  slideRels.items.push({
    id: 'rId901',
    type: COMMENTS_REL,
    target: COMMENTS,
    targetMode: 'Internal',
  });
  pkg.setRels(partName('/ppt/slides/slide1.xml'), slideRels);
  // Round-trip once so every later read comes off bytes, not off the
  // in-memory parts this helper just pushed in.
  return loadPresentation(await savePresentation(pres));
};

const commentsXmlOf = (pres: PresentationData): string => {
  const part = _internalPackageOf(pres).parts.find((p) => p.name === COMMENTS);
  if (!part) throw new Error('expected the modern comment part');
  return decode(part.data);
};

const reload = async (pres: PresentationData): Promise<PresentationData> =>
  loadPresentation(await savePresentation(pres));

describe('modern comments', () => {
  it('reads threads, replies, authors and resolved state', async () => {
    const pres = await withModernComments();
    const comments = getSlideComments(getSlides(pres)[0]!);
    expect(comments.map(getCommentText)).toEqual([
      'Tighten this headline.',
      '承知しました。',
      'This chart needs a source line.',
    ]);
    expect(comments.map(getCommentFormat)).toEqual(['modern', 'modern', 'modern']);
    expect(comments.map(getCommentStatus)).toEqual(['resolved', 'active', 'active']);
    expect(comments.map((c) => getCommentAuthor(c).name)).toEqual([
      'Reviewer',
      '校閲 二郎',
      '校閲 二郎',
    ]);
    expect(getCommentAuthor(comments[1]!).initials).toBe('校');
    // A reply answers the thread it sits in; the threads themselves answer
    // nothing.
    expect(getCommentParent(comments[0]!)).toBeNull();
    expect(getCommentText(getCommentParent(comments[1]!)!)).toBe('Tighten this headline.');
    expect(getCommentParent(comments[2]!)).toBeNull();
    expect(getCommentDate(comments[0]!)).toBe('2026-05-15T12:00:00.000');
  });

  it('resolves a thread and reopens it, in the file', async () => {
    const pres = await withModernComments();
    const open = getSlideComments(getSlides(pres)[0]!)[2]!;
    setCommentStatus(open, 'resolved');
    const afterResolve = await reload(pres);
    expect(getSlideComments(getSlides(afterResolve)[0]!).map(getCommentStatus)).toEqual([
      'resolved',
      'active',
      'resolved',
    ]);

    const again = getSlideComments(getSlides(afterResolve)[0]!)[0]!;
    setCommentStatus(again, 'active');
    const afterReopen = await reload(afterResolve);
    expect(getSlideComments(getSlides(afterReopen)[0]!).map(getCommentStatus)).toEqual([
      'active',
      'active',
      'resolved',
    ]);
  });

  it('leaves everything it does not model exactly where it was', async () => {
    const pres = await withModernComments();
    const slide = getSlides(pres)[0]!;
    const [thread, reply, anchored] = getSlideComments(slide);
    setCommentStatus(thread!, 'active');
    setCommentText(thread!, 'Tighten this headline, and say who by.');
    setCommentText(reply!, '直しました。');
    const saved = commentsXmlOf(await reload(pres));
    // The reaction, its extension list, and the shape anchor of the thread
    // nobody touched.
    expect(saved).toContain('<p223:reaction authorId="' + SECOND + '" type="like"/>');
    expect(saved).toContain('uri="{9E7C0F2E-0000-0000-0000-00000000FEED}"');
    expect(saved).toContain('<ac:deMk id="7" idx="0"/>');
    // And the pin, and the slide moniker.
    expect(saved).toContain('<p188:pos x="1270000" y="635000"/>');
    expect(saved).toContain('<pc:sldMk cId="3141592653" sldId="256"/>');
    expect(getCommentText(anchored!)).toBe('This chart needs a source line.');

    const back = getSlideComments(getSlides(await reload(pres))[0]!);
    expect(back.map(getCommentText)).toEqual([
      'Tighten this headline, and say who by.',
      '直しました。',
      'This chart needs a source line.',
    ]);
    expect(back.map(getCommentStatus)).toEqual(['active', 'active', 'active']);
  });

  it('replies inside the thread it answers, as an author already known', async () => {
    const pres = await withModernComments();
    const slide = getSlides(pres)[0]!;
    const thread = getSlideComments(slide)[0]!;
    addSlideComment(slide, {
      author: { name: 'Reviewer', initials: 'R' },
      text: 'One more thing.',
      replyTo: thread,
      date: new Date('2026-05-17T08:00:00.000Z'),
    });
    const after = await reload(pres);
    const comments = getSlideComments(getSlides(after)[0]!);
    expect(comments.map(getCommentText)).toEqual([
      'Tighten this headline.',
      '承知しました。',
      'One more thing.',
      'This chart needs a source line.',
    ]);
    expect(getCommentText(getCommentParent(comments[2]!)!)).toBe('Tighten this headline.');
    // The author was already in the list, so no second record was made for
    // the same person.
    const authors = _internalPackageOf(after).parts.find((p) => p.name === AUTHORS);
    expect(decode(authors!.data).match(/<p188:author /g)).toHaveLength(2);
  });

  it('starts a new thread in the format the slide already uses', async () => {
    const pres = await withModernComments();
    const slide = getSlides(pres)[0]!;
    const added = addSlideComment(slide, {
      author: { name: 'Newcomer', initials: 'N' },
      text: 'Whole-slide note.',
      date: new Date('2026-05-18T08:00:00.000Z'),
    });
    expect(getCommentFormat(added)).toBe('modern');
    expect(getCommentStatus(added)).toBe('active');
    const after = await reload(pres);
    const saved = commentsXmlOf(after);
    // The anchor is the one the slide's own thread uses, not a guess: the
    // fixture had one slide moniker list, and now there are two alike.
    expect(saved.match(/<pc:sldMkLst>/g)).toHaveLength(2);
    expect(saved.match(/<pc:sldMk cId="3141592653" sldId="256"\/>/g)).toHaveLength(3);
    const authors = decode(_internalPackageOf(after).parts.find((p) => p.name === AUTHORS)!.data);
    expect(authors).toContain('name="Newcomer"');
    // No identity provider is being claimed on this person's behalf.
    expect(authors).toContain('userId="" providerId=""');
    // The legacy list was not started as a side effect.
    expect(
      _internalPackageOf(after).parts.some((p) => p.name === '/ppt/comments/comment1.xml'),
    ).toBe(false);
    expect(getSlideComments(getSlides(after)[0]!).map(getCommentText)).toContain(
      'Whole-slide note.',
    );
  });

  it('removes one reply, or a whole thread with its replies', async () => {
    const pres = await withModernComments();
    removeSlideComment(getSlideComments(getSlides(pres)[0]!)[1]!);
    const afterReply = await reload(pres);
    expect(getSlideComments(getSlides(afterReply)[0]!).map(getCommentText)).toEqual([
      'Tighten this headline.',
      'This chart needs a source line.',
    ]);
    // An empty reply list is not left behind.
    expect(commentsXmlOf(afterReply)).not.toContain('replyLst');

    removeSlideComment(getSlideComments(getSlides(afterReply)[0]!)[0]!);
    const afterThread = await reload(afterReply);
    expect(getSlideComments(getSlides(afterThread)[0]!).map(getCommentText)).toEqual([
      'This chart needs a source line.',
    ]);
  });

  it('drops the part and its relationship once the last thread goes', async () => {
    const pres = await withModernComments();
    const slide = getSlides(pres)[0]!;
    expect(clearSlideComments(slide)).toBe(3);
    const after = await reload(pres);
    const pkg = _internalPackageOf(after);
    expect(pkg.parts.some((p) => p.name === COMMENTS)).toBe(false);
    expect(
      pkg.getRels(partName('/ppt/slides/slide1.xml'))?.items.some((r) => r.type === COMMENTS_REL),
    ).toBe(false);
    expect(getSlideComments(getSlides(after)[0]!)).toEqual([]);
  });

  it('refuses to record a status on a comment that has nowhere to keep one', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const legacy = addSlideComment(slide, {
      author: { name: 'Reviewer', initials: 'R' },
      text: 'An ECMA-376 comment.',
    });
    expect(getCommentFormat(legacy)).toBe('legacy');
    expect(getCommentStatus(legacy)).toBeNull();
    expect(() => setCommentStatus(legacy, 'resolved')).toThrow(/cannot record a status/);
  });

  it('keeps a deck that has both formats reading and editing as two lists', async () => {
    const pres = await withModernComments();
    const slide = getSlides(pres)[0]!;
    // A slide that already has modern comments keeps taking modern ones, so
    // the legacy list here is built deliberately, the way an older tool would.
    const pkg = _internalPackageOf(pres);
    pkg.addPart(
      partName('/ppt/commentAuthors.xml'),
      'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml',
      encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cmAuthor id="0" name="Older Tool" initials="O" lastIdx="1" clrIdx="0"/></p:cmAuthorLst>`,
      ),
    );
    pkg.addPart(
      partName('/ppt/comments/comment1.xml'),
      'application/vnd.openxmlformats-officedocument.presentationml.comments+xml',
      encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cm authorId="0" idx="1" dt="2026-04-01T09:00:00Z"><p:pos x="0" y="0"/><p:text>From an older tool.</p:text></p:cm></p:cmLst>`,
      ),
    );
    const presRels = pkg.getRels(partName('/ppt/presentation.xml'))!;
    presRels.items.push({
      id: 'rId902',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors',
      target: '/ppt/commentAuthors.xml',
      targetMode: 'Internal',
    });
    pkg.setRels(partName('/ppt/presentation.xml'), presRels);
    const slideRels = pkg.getRels(partName('/ppt/slides/slide1.xml'))!;
    slideRels.items.push({
      id: 'rId903',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
      target: '/ppt/comments/comment1.xml',
      targetMode: 'Internal',
    });
    pkg.setRels(partName('/ppt/slides/slide1.xml'), slideRels);

    const both = await reload(pres);
    const comments = getSlideComments(getSlides(both)[0]!);
    expect(comments.map(getCommentFormat)).toEqual(['legacy', 'modern', 'modern', 'modern']);
    expect(getCommentText(comments[0]!)).toBe('From an older tool.');
    // Editing one does not disturb the other.
    setCommentText(comments[0]!, 'Edited in place.');
    setCommentStatus(comments[1]!, 'active');
    const after = await reload(both);
    const back = getSlideComments(getSlides(after)[0]!);
    expect(back.map(getCommentText)).toEqual([
      'Edited in place.',
      'Tighten this headline.',
      '承知しました。',
      'This chart needs a source line.',
    ]);
    expect(back.map(getCommentStatus)).toEqual([null, 'active', 'active', 'active']);
    void slide;
  });
});
