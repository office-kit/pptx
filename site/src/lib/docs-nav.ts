// Single source of truth for the docs sidebar + sitemap-style listings
// (also used by /llms.txt to list every documentation page).

export type DocLink = {
  /** Absolute path on the site, no trailing slash */
  href: string;
  /** Title shown in the sidebar / llms.txt */
  title: string;
  /** Short description used by llms.txt */
  description: string;
};

export type DocSection = {
  title: string;
  links: DocLink[];
};

export const docSections: DocSection[] = [
  {
    title: 'Getting started',
    links: [
      {
        href: '/docs/authoring',
        title: 'Create slides with AI',
        description: 'Create a TSX project, edit with Claude Code, preview, and export a PPTX.',
      },
      {
        href: '/docs/tsx',
        title: 'TSX elements and templates',
        description: 'Typed elements, template references, units, and the Raw escape hatch.',
      },
      {
        href: '/docs/install',
        title: 'Core library install',
        description: 'Add @office-kit/pptx to a Node or browser project.',
      },
      {
        href: '/docs/getting-started',
        title: 'Core library workflows',
        description: 'Read, edit, and write your first .pptx presentation.',
      },
      {
        href: '/docs/recipes',
        title: 'Recipes',
        description:
          'Working code for the most common tasks — load / edit / build / charts / images / notes / validate.',
      },
      {
        href: '/docs/cheatsheet',
        title: 'Cheatsheet',
        description:
          'One-page lookup of task → exact functions to import + call. The shortest path from "I want to do X" to working code.',
      },
    ],
  },
  {
    title: 'Reference',
    links: [
      {
        href: '/docs/api',
        title: 'API overview',
        description: 'Public surface: free-function API + units + the Node convenience entry.',
      },
    ],
  },
];

export const allDocLinks: DocLink[] = docSections.flatMap((s) => s.links);
