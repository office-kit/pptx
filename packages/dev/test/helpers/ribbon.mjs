export async function openArrange(editor, locale = 'en') {
  await editor.getByRole('tab', { name: locale === 'ja' ? 'ホーム' : 'Home', exact: true }).click();
  // A narrow Home ribbon collapses Drawing into one button, as the reference desktop app does.
  const drawing = editor
    .getByRole('tabpanel')
    .getByRole('button', { name: locale === 'ja' ? '図形描画' : 'Drawing', exact: true });
  if ((await drawing.isVisible()) && (await drawing.getAttribute('aria-expanded')) !== 'true') {
    await drawing.click();
  }
  await editor
    .getByRole('button', { name: locale === 'ja' ? '配置' : 'Arrange', exact: true })
    .click();
}
