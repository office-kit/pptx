/** Select UTF-16 offsets independently of how the editor splits its styled runs. */
export async function installRichTextSelection(page) {
  await page.addInitScript(() => {
    window.selectEditorText = (node, start, end = start) => {
      if (node instanceof HTMLTextAreaElement || node instanceof HTMLInputElement) {
        node.setSelectionRange(start, end);
        return;
      }
      const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT);
      let current = walker.nextNode();
      let offset = 0;
      const points = [];
      for (const target of [start, end]) {
        while (current && offset + current.textContent.length < target) {
          offset += current.textContent.length;
          current = walker.nextNode();
        }
        points.push(current ? [current, target - offset] : [node, node.childNodes.length]);
      }
      const range = document.createRange();
      range.setStart(...points[0]);
      range.setEnd(...points[1]);
      const selection = document.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      document.dispatchEvent(new Event('selectionchange'));
    };
  });
}
