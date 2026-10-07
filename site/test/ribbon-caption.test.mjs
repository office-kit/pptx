import assert from 'node:assert/strict';
import test from 'node:test';
import { captionLines } from '../src/lib/editor/ribbon/caption.ts';

test('English captions break at the balancing space, preferring the later word', () => {
  assert.equal(captionLines('Table'), 'Table');
  assert.equal(captionLines('New Slide'), 'New\nSlide');
  assert.equal(captionLines('Header & Footer'), 'Header &\nFooter');
  assert.equal(captionLines('Date & Time'), 'Date &\nTime');
  assert.equal(captionLines('Ink to Text'), 'Ink to\nText');
  assert.equal(captionLines('Draw with Trackpad'), 'Draw with\nTrackpad');
});

test('Japanese captions break after a particle or where the script changes', () => {
  assert.equal(captionLines('新しいスライド'), '新しい\nスライド');
  assert.equal(captionLines('ヘッダーとフッター'), 'ヘッダーと\nフッター');
  assert.equal(captionLines('インクをテキストに変換'), 'インクを\nテキストに変換');
  assert.equal(captionLines('背景のスタイル'), '背景の\nスタイル');
  assert.equal(captionLines('なげなわ選択'), 'なげなわ\n選択');
  assert.equal(captionLines('スクリーンショット'), 'スクリーンショット');
  assert.equal(captionLines('日付と時刻'), '日付と時刻');
});
