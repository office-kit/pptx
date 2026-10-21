const ALLOWED_PROTOCOLS = new Set(['http:', 'https:', 'mailto:']);

// A deck's hyperlink targets are untrusted, and the SVG may be inlined into a
// host page, so only web, mail and in-document links become `<a href>`.
// Parsing with the WHATWG URL parser sees the scheme exactly as the browser
// will (case, surrounding spaces and embedded tabs/newlines all normalised).
export const safeLinkHref = (href: string): string | undefined => {
  if (href.startsWith('#')) return href;
  try {
    return ALLOWED_PROTOCOLS.has(new URL(href).protocol) ? href : undefined;
  } catch {
    return undefined;
  }
};
