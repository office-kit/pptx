// Normal view shows the notes pane by default, as the reference desktop app does. Open it
// only when hidden (the Notes button toggles), then focus it as the button
// does when it opens the pane.
export async function openNotes(editor) {
  const button = editor.getByRole('button', { name: 'Notes', exact: true });
  if ((await button.getAttribute('aria-pressed')) !== 'true') await button.click();
  else await editor.getByRole('textbox', { name: 'Notes content', exact: true }).focus();
}
