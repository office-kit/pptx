import { describe, expect, it } from 'vitest';
import {
  editFrameCropGeometry,
  editPictureCropGeometry,
  getPictureCropGeometry,
  resetCropGeometry,
} from '../packages/editor/src/core/crop-geometry.ts';

const frame = { x: 2.54, y: 2.54, w: 10.16, h: 5.715 };
const close = (value: number, expected: number) => expect(value).toBeCloseTo(expected, 5);

describe('editor crop geometry', () => {
  it('maps a centered crop to native picture fields', () => {
    const g = getPictureCropGeometry(frame);
    close(g.pictureWidth, 10.16);
    close(g.pictureHeight, 5.715);
    close(g.offsetX, 0);
    close(g.offsetY, 0);
    close(g.cropWidth, 10.16);
    close(g.cropHeight, 5.715);
    close(g.cropLeft, frame.x);
    close(g.cropTop, frame.y);
  });

  it('uses center-relative offsets and preserves the frame when editing the picture', () => {
    const edited = editPictureCropGeometry(
      frame,
      { left: 0.07667, right: 0.07667 },
      {
        pictureWidth: 12,
        offsetX: 1,
      },
    );
    expect(edited.frame).toEqual(frame);
    close(edited.crop.left, -1 / 150);
    close(edited.crop.right, 0.16);
    close(getPictureCropGeometry(frame, edited.crop).offsetX, 1);
    close(getPictureCropGeometry(frame, edited.crop).cropLeft, frame.x);
    close(getPictureCropGeometry(frame, edited.crop).cropTop, frame.y);
  });

  it('matches the native crop coordinates after a picture-width edit', () => {
    const g = getPictureCropGeometry(frame, { left: 0.07667, right: 0.07667 });
    close(g.pictureWidth, 12.00009449);
    close(g.offsetX, 0);
    close(g.cropLeft, 2.54);
    close(g.cropTop, 2.54);
  });

  it('keeps the absolute picture rectangle when editing frame bounds', () => {
    const first = editPictureCropGeometry(frame, null, { pictureWidth: 12, offsetX: 1 });
    const next = editFrameCropGeometry(first.frame, first.crop, { w: 8, x: 3 });
    close(next.crop.left, 0.0316666667);
    close(next.crop.right, 0.3016666667);
    close(getPictureCropGeometry(next.frame, next.crop).pictureWidth, 12);
    close(getPictureCropGeometry(next.frame, next.crop).offsetX, 1.62);
  });

  it('resets to the picture rectangle and clears the crop', () => {
    const current = editPictureCropGeometry(frame, null, { pictureWidth: 12, offsetX: 1 });
    const reset = resetCropGeometry(current.frame, current.crop);
    close(reset.frame.x, 2.62);
    close(reset.frame.y, 2.54);
    close(reset.frame.w, 12);
    close(reset.frame.h, 5.715);
    expect(reset.crop).toBeNull();
  });

  it('supports negative crop sides and rejects invalid geometry', () => {
    const g = getPictureCropGeometry(frame, { left: -0.1, top: -0.2 });
    close(g.cropLeft, frame.x);
    close(g.cropTop, frame.y);
    expect(() => getPictureCropGeometry({ ...frame, w: 0 })).toThrow(/positive/);
    expect(() => getPictureCropGeometry(frame, { left: 1, right: 0 })).toThrow(/positive/);
    expect(() => editPictureCropGeometry(frame, null, { offsetX: Number.NaN })).toThrow(/finite/);
  });
});
