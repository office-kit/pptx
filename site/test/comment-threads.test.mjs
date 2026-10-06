import assert from 'node:assert/strict';
import test from 'node:test';
import { orderCommentThreads } from '../src/lib/editor/core/comment-threads.ts';

test('groups interleaved replies after their parents without changing save order', () => {
  const comments = [
    { id: 0, parent: null },
    { id: 1, parent: null },
    { id: 2, parent: 0 },
    { id: 3, parent: 2 },
    { id: 4, parent: 0 },
  ];
  assert.deepEqual(
    orderCommentThreads(comments).map((comment) => comment.id),
    [0, 2, 3, 4, 1],
  );
  assert.deepEqual(
    comments.map((comment) => comment.id),
    [0, 1, 2, 3, 4],
  );
});

test('keeps orphaned and cyclic imported threads visible exactly once', () => {
  const comments = [
    { id: 0, parent: 99 },
    { id: 1, parent: 2 },
    { id: 2, parent: 1 },
    { id: 3, parent: 3 },
    { id: 4, parent: 0 },
  ];
  assert.deepEqual(
    orderCommentThreads(comments).map((comment) => comment.id),
    [0, 4, 1, 2, 3],
  );
});

test('handles deep reply chains without recursive stack growth', () => {
  const comments = Array.from({ length: 20000 }, (_, id) => ({
    id,
    parent: id === 0 ? null : id - 1,
  }));
  assert.deepEqual(orderCommentThreads(comments), comments);
});
