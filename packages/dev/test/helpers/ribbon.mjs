export async function openArrange(editor, locale = 'en') {
  await editor.getByRole('tab', { name: locale === 'ja' ? 'ホーム' : 'Home', exact: true }).click();
  const drawing = editor
    .getByRole('toolbar', {
      name: locale === 'ja' ? 'ホームリボンのグループ' : 'Home ribbon groups',
      exact: true,
    })
    .getByRole('button', { name: locale === 'ja' ? '図形描画' : 'Drawing', exact: true });
  if ((await drawing.isVisible()) && (await drawing.getAttribute('aria-expanded')) !== 'true') {
    await drawing.click();
  }
  await editor
    .getByRole('button', { name: locale === 'ja' ? '配置' : 'Arrange', exact: true })
    .click();
}
