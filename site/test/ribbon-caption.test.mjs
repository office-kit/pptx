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

test('English captions balance by rendered width, and no line starts with &', () => {
  // Widths at 11 pt in San Francisco: "Insert Row" has more characters than
  // "Row Above" but renders narrower.
  const widths = {
    Insert: 30.1,
    'Row Above': 58,
    'Insert Row': 55.3,
    Above: 32.8,
    'Insert Column': 73.2,
    Right: 27.4,
    'Column Right': 70.5,
  };
  const measure = (text) => widths[text];
  assert.equal(captionLines('Insert Row Above', measure), 'Insert Row\nAbove');
  assert.equal(captionLines('Insert Column Right', measure), 'Insert\nColumn Right');
  // "Header" / "& Footer" would be narrower, but "&" never starts a line.
  assert.equal(
    captionLines('Header & Footer', (text) => text.length * 6),
    'Header &\nFooter',
  );
});

test('へ is a particle too, so no line starts with it', () => {
  assert.equal(captionLines('最前面へ移動'), '最前面へ\n移動');
  assert.equal(captionLines('最背面へ移動'), '最背面へ\n移動');
});
