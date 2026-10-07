/** Keep conversation order separate from the order in which drafts must be saved. */
export function orderCommentThreads<T extends { id: number; parent: number | null }>(
  comments: readonly T[],
): T[] {
  const ids = new Set(comments.map((comment) => comment.id));
  const children = new Map<number, T[]>();
  const roots: T[] = [];
  for (const comment of comments) {
    if (comment.parent === null || !ids.has(comment.parent)) roots.push(comment);
    else {
      const siblings = children.get(comment.parent) ?? [];
      siblings.push(comment);
      children.set(comment.parent, siblings);
    }
  }
  const result: T[] = [];
  const visited = new Set<number>();
  function visit(root: T) {
    const stack = [root];
    while (stack.length) {
      const comment = stack.pop()!;
      if (visited.has(comment.id)) continue;
      visited.add(comment.id);
      result.push(comment);
      const replies = children.get(comment.id) ?? [];
      for (let index = replies.length - 1; index >= 0; index--) stack.push(replies[index]!);
    }
  }
  for (const root of roots) visit(root);
  // Imported cycles have no root. Still expose every comment so users can edit or delete it.
  for (const comment of comments) if (!visited.has(comment.id)) visit(comment);
  return result;
}
