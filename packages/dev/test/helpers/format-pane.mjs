// Like Mac PowerPoint, every Format pane section starts collapsed; tests open
// the ones they use by their header, exactly as a user would.
const escape = (text) => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

export async function expandFormatSections(editor, ...names) {
  for (const name of names) {
    const section = editor
      .locator('details.pane-section:visible')
      .filter({
        has: editor.locator(':scope > summary', { hasText: new RegExp(`^${escape(name)}$`) }),
      })
      .first();
    if (!(await section.evaluate((node) => node.open)))
      await section.locator(':scope > summary').click();
  }
}
