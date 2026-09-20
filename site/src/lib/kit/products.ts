// The Office Kit family. This registry is the only place that knows where each
// product's site lives, so the header switcher, the footer, and the landing
// page's family grid all stay in step.
//
// Today each product deploys its own site from its own repo to
// `office-kit.github.io/<id>/`. When the three merge into one site, swap
// `href` for an in-site path here and nothing else has to change.

export type ProductId = 'pptx' | 'xlsx' | 'docx';

export type Product = {
  id: ProductId;
  /** npm package name. */
  pkg: string;
  /** The Office application whose files this product reads and writes. */
  app: string;
  /** One line on what the product covers, shown in the family listings. */
  summary: string;
  href: string;
  repo: string;
};

const SITE_ORIGIN = 'https://office-kit.github.io';
const ORG = 'https://github.com/office-kit';

export const products: readonly Product[] = [
  {
    id: 'pptx',
    pkg: '@office-kit/pptx',
    app: 'PowerPoint',
    summary: 'Slides, charts, tables, themes, notes, and animations.',
    href: `${SITE_ORIGIN}/pptx/`,
    repo: `${ORG}/pptx`,
  },
  {
    id: 'xlsx',
    pkg: '@office-kit/xlsx',
    app: 'Excel',
    summary: 'Workbooks, formulas, styles, and streaming for large sheets.',
    href: `${SITE_ORIGIN}/xlsx/`,
    repo: `${ORG}/xlsx`,
  },
  {
    id: 'docx',
    pkg: '@office-kit/docx',
    app: 'Word',
    summary: 'Paragraphs, lists, tables, page setup, and template editing.',
    href: `${SITE_ORIGIN}/docx/`,
    repo: `${ORG}/docx`,
  },
];

/** The product this build of the site documents. */
export const CURRENT_PRODUCT: ProductId = 'pptx';

export const currentProduct: Product = products.find((p) => p.id === CURRENT_PRODUCT)!;

export const ORG_URL = ORG;
