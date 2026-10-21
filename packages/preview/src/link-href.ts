const ALLOWED_PROTOCOLS = new Set(['http:', 'https:', 'mailto:']);

// Hyperlink targets are untrusted and the SVG may be inlined into a host page,
// so only web, mail and in-document links become `<a href>`.
export const safeLinkHref = (href: string): string | undefined => {
  if (href.startsWith('#')) return href;
  try {
    return ALLOWED_PROTOCOLS.has(new URL(href).protocol) ? href : undefined;
  } catch {
    return undefined;
  }
};
