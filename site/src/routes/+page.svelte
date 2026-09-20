<script lang="ts">
  import { base } from '$app/paths';
  import { FamilyGrid, InstallCommand, getProduct } from '@office-kit/site-kit';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();

  const currentProduct = getProduct('pptx');

  const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  const DOWNLOAD_NAME = 'office-kit-demo.pptx';

  let download = $state<{ state: 'idle' | 'working' | 'done' | 'failed'; note: string }>({
    state: 'idle',
    note: '',
  });

  // Runs the exact function shown in the code panel, in the visitor's browser.
  // The library loads on click so the landing page itself ships none of it.
  async function downloadDeck(): Promise<void> {
    download = { state: 'working', note: '' };
    try {
      const [{ savePresentation }, { buildHeroDeck }] = await Promise.all([
        import('@office-kit/pptx'),
        import('$lib/examples/hero-deck'),
      ]);
      const bytes = await savePresentation(buildHeroDeck());
      const url = URL.createObjectURL(new Blob([bytes.slice()], { type: PPTX_MIME }));
      const a = document.createElement('a');
      a.href = url;
      a.download = DOWNLOAD_NAME;
      a.click();
      URL.revokeObjectURL(url);
      download = {
        state: 'done',
        note: `Saved ${DOWNLOAD_NAME} (${Math.round(bytes.byteLength / 1024)} KB), built in this tab.`,
      };
    } catch (err) {
      console.error(err);
      download = {
        state: 'failed',
        note: 'The deck could not be built in this browser. The console has the error.',
      };
    }
  }

  const paths = [
    {
      title: 'Edit a file you already have',
      body: 'Any .pptx made in PowerPoint, Keynote, or Google Slides is a template. Fill {{tokens}}, swap a picture and keep its crop, add slides from the deck’s own layouts. Parts the library does not model, such as SmartArt, OLE objects, and video, are carried through unchanged.',
      code: 'loadPresentation(bytes)',
      href: '/docs/getting-started',
      link: 'Read the template guide',
    },
    {
      title: 'Create slides with AI',
      body: 'Install the Claude Code skill and describe your presentation. Claude handles setup, TSX authoring, live preview, and PowerPoint export. Review the slides and ask for changes.',
      code: 'deck.tsx → preview → deck.pptx',
      href: '/docs/authoring',
      link: 'Create your first presentation',
    },
    {
      title: 'See the slide before you send it',
      body: 'The companion @office-kit/pptx-preview package draws a slide as SVG in the browser and as PNG in Node, with no headless Office. It also reports text that overflows its box, which is how an automated pipeline catches a broken slide.',
      code: 'renderSlideToSvg(pres, slide)',
      href: '/playground',
      link: 'Try it in the playground',
    },
  ];

  const proof = [
    {
      claim: 'Every sample deck passes Microsoft’s Open XML SDK validator.',
      how: 'A CI job runs OpenXmlValidator over the generated samples and fails the build on the first error. Emitted XML is also checked against the ECMA-376 schemas.',
    },
    {
      claim: 'What the library does not understand, it does not touch.',
      how: 'Unknown parts and extensions are preserved on load → save, so a deck that went in with SmartArt comes out with SmartArt.',
    },
    {
      claim: 'You ship only the functions you import.',
      how: 'Side-effect-free exports let bundlers remove unused functions. CI checks tree-shaking and bundle-size limits.',
    },
    {
      claim: 'One runtime dependency.',
      how: 'fflate, for ZIP. XML reading and writing is hand-rolled, and nothing on the hot path touches fs, Buffer, or zlib.',
    },
  ];

  type Cell = { text: string; tone?: 'yes' | 'no' | 'part' };
  const comparison: Array<{ topic: string; ours: Cell; theirs: Cell }> = [
    {
      topic: 'Open and edit an existing .pptx',
      ours: { text: 'Yes. Load, change, save; unknown parts are preserved', tone: 'yes' },
      theirs: { text: 'No. It creates new files only', tone: 'no' },
    },
    {
      topic: 'Templates',
      ours: { text: 'Any .pptx a designer made in PowerPoint', tone: 'yes' },
      theirs: { text: 'Slide masters defined in code', tone: 'part' },
    },
    {
      topic: 'Read back what is in a deck',
      ours: { text: 'Every setter has a getter, plus deck-wide queries', tone: 'yes' },
      theirs: { text: 'No. The API is write-only', tone: 'no' },
    },
    {
      topic: 'API shape',
      ours: { text: 'Tree-shakeable functions, ESM only' },
      theirs: { text: 'One class with methods; ESM, CommonJS, and a script-tag bundle' },
    },
    {
      topic: 'Transitions and animations',
      ours: { text: 'Slide transitions; four entrance and exit presets', tone: 'yes' },
      theirs: { text: 'No', tone: 'no' },
    },
    {
      topic: 'Comments',
      ours: { text: 'Yes', tone: 'yes' },
      theirs: { text: 'No (speaker notes only)', tone: 'no' },
    },
    {
      topic: 'Chart types you can author',
      ours: {
        text: 'Bar, column, line, pie, doughnut, area; combos, secondary axis, trendlines',
        tone: 'part',
      },
      theirs: { text: 'Those plus scatter, bubble, radar, and 3D bar', tone: 'yes' },
    },
    {
      topic: 'Audio, video, YouTube embeds',
      ours: { text: 'Not yet. Media already in a deck is preserved', tone: 'no' },
      theirs: { text: 'Yes', tone: 'yes' },
    },
    {
      topic: 'HTML table to slides',
      ours: { text: 'No', tone: 'no' },
      theirs: { text: 'Yes, with automatic paging', tone: 'yes' },
    },
    {
      topic: 'Render a slide to an image',
      ours: { text: 'SVG and PNG, via the companion preview package', tone: 'yes' },
      theirs: { text: 'No', tone: 'no' },
    },
  ];

  const capabilities = [
    { area: 'Slides', items: 'Add, remove, move, duplicate, hide. Layouts, sections, sizes, backgrounds.' },
    { area: 'Text', items: 'Paragraphs, runs, bullets, spacing, autofit. Latin, East Asian, and complex-script fonts.' },
    { area: 'Shapes', items: '180+ presets, lines, groups, fills (solid, gradient, pattern, picture), strokes, shadows, glows, z-order.' },
    { area: 'Tables', items: 'Rows, columns, merges, borders, fills, per-cell text, margins, and direction.' },
    { area: 'Charts', items: 'Six kinds plus combos, secondary axis, data labels, trendlines. The workbook behind “Edit data” is generated for you.' },
    { area: 'Pictures', items: 'PNG, JPEG, GIF, BMP, TIFF, WebP, SVG. Crop, opacity, brightness, contrast, in-place swap.' },
    { area: 'Review', items: 'Speaker notes, comments, hyperlinks, click actions, with deck-wide find and replace.' },
    { area: 'Deck', items: 'Theme colors and fonts, document properties, thumbnail, package inspection, validatePresentation().' },
  ];

  const notYet = [
    'New themes, masters, and layouts from scratch',
    'SmartArt and OLE authoring (both are preserved on round trip)',
    'Audio and video authoring',
    'Scatter, radar, and bubble chart authoring (they read and render)',
    'Encrypted files',
  ];
</script>

<svelte:head>
  <title>@office-kit/pptx: read, edit, and write PowerPoint files in TypeScript</title>
</svelte:head>

<section class="band hero">
  <div class="frame hero-inner">
    <h1>Read, edit, and write PowerPoint files in TypeScript</h1>
    <p class="lede">
      Open any .pptx or start from an empty deck. Change slides, charts, and tables through typed
      functions, then save a file that PowerPoint, Keynote, Google Slides, and LibreOffice all open
      cleanly. It runs in Node and in the browser.
    </p>
    <div class="cta">
      <a href="{base}/docs/authoring" class="btn primary">Create slides with AI</a>
      <a href="{base}/playground" class="btn">Open the playground</a>
      <InstallCommand pkg={currentProduct.pkg} />
    </div>
  </div>
</section>

<section class="band stage" aria-labelledby="stage-title">
  <div class="frame stage-inner">
    <h2 id="stage-title" class="visually-hidden">A slide and the code that built it</h2>
    <div class="stage-grid">
      <figure class="code-pane">
        <figcaption>hero-deck.ts</figcaption>
        <!-- A scrollable region must be focusable, or keyboard users cannot scroll it. -->
        <!-- svelte-ignore a11y_no_noninteractive_tabindex -->
        <div class="code-scroll" tabindex="0" role="region" aria-label="Source of hero-deck.ts">
          {@html data.heroCode}
        </div>
      </figure>
      <figure class="slide-pane">
        <div class="slide" role="img" aria-label="The rendered slide: a column chart titled Revenue grew 2.5× in four quarters">
          {@html data.slideSvg}
        </div>
      </figure>
    </div>
    <div class="stage-foot">
      <p>
        That slide is real output. The code shown with it built the deck, and
        <code>@office-kit/pptx-preview</code> drew it as SVG. Download it and open it in PowerPoint:
        the chart is editable, data and all.
      </p>
      <div class="stage-action">
        <button
          type="button"
          class="btn stage-btn"
          onclick={downloadDeck}
          disabled={download.state === 'working'}
        >
          {download.state === 'working' ? 'Building the deck…' : 'Download the .pptx'}
        </button>
        <p class="stage-note" aria-live="polite">{download.note}</p>
      </div>
    </div>
  </div>
</section>

<section class="band">
  <div class="frame">
    <ul class="paths">
      {#each paths as p (p.title)}
        <li>
          <code class="path-code">{p.code}</code>
          <h2>{p.title}</h2>
          <p>{p.body}</p>
          <a href="{base}{p.href}">{p.link}</a>
        </li>
      {/each}
    </ul>
  </div>
</section>

<section class="band template">
  <div class="frame template-inner">
    <div class="template-text">
      <h2>A designer makes the template. Your code fills it in.</h2>
      <p>
        Most decks that matter already exist as a .pptx someone cares about: the brand template, last
        quarter’s review, the certificate with the right fonts. Load it, replace what changes, and
        everything else stays exactly as it was designed.
      </p>
      <p>
        Every write has a matching read, so the same API that fills a template can also audit one:
        list the layouts, find every hyperlink, or check which slides still have an empty title.
      </p>
    </div>
    <figure class="template-code">
      <figcaption>template-fill.ts</figcaption>
      {@html data.templateCode}
    </figure>
  </div>
</section>

<section class="band proof" aria-labelledby="proof-title">
  <div class="frame proof-inner">
    <h2 id="proof-title">Valid files, checked by machines</h2>
    <p class="proof-lede">
      A deck that opens in PowerPoint but breaks Keynote is a bug here. So is one that opens
      everywhere and still fails schema validation.
    </p>
    <dl>
      {#each proof as item (item.claim)}
        <div class="proof-row">
          <dt>{item.claim}</dt>
          <dd>{item.how}</dd>
        </div>
      {/each}
    </dl>
  </div>
</section>

<section class="band" aria-labelledby="compare-title">
  <div class="frame compare">
    <div class="section-head">
      <h2 id="compare-title">How it differs from PptxGenJS</h2>
      <p>
        PptxGenJS is the established way to generate a deck in JavaScript, and it is good at that.
        The difference is direction: PptxGenJS writes new files, while this library reads, edits,
        and writes. Compared against PptxGenJS 4.0.
      </p>
    </div>
    <div class="table-scroll">
      <table>
        <thead>
          <tr>
            <th scope="col"><span class="visually-hidden">Capability</span></th>
            <th scope="col">@office-kit/pptx</th>
            <th scope="col">PptxGenJS</th>
          </tr>
        </thead>
        <tbody>
          {#each comparison as row (row.topic)}
            <tr>
              <th scope="row">{row.topic}</th>
              <td data-tone={row.ours.tone}>{row.ours.text}</td>
              <td data-tone={row.theirs.tone}>{row.theirs.text}</td>
            </tr>
          {/each}
        </tbody>
      </table>
    </div>
    <p class="compare-advice">
      Pick PptxGenJS if you only ever generate new decks and need video, scatter charts, or a
      script-tag build. Pick this library if a template, an existing deck, or a validation
      requirement is involved.
    </p>
  </div>
</section>

<section class="band" aria-labelledby="caps-title">
  <div class="frame caps">
    <div class="section-head">
      <h2 id="caps-title">What you can build today</h2>
      <p>
        The library is pre-1.0 and says so. This is what works now, and what does not yet.
        <a href="{base}/api">The API reference</a> lists every function.
      </p>
    </div>
    <dl class="caps-grid">
      {#each capabilities as c (c.area)}
        <div>
          <dt>{c.area}</dt>
          <dd>{c.items}</dd>
        </div>
      {/each}
    </dl>
    <div class="not-yet">
      <h3>Not yet</h3>
      <ul>
        {#each notYet as item (item)}
          <li>{item}</li>
        {/each}
      </ul>
    </div>
  </div>
</section>

<section class="band agents">
  <div class="frame agents-inner">
    <div>
      <h2>Written to be driven by AI agents too</h2>
      <p>
        Decks are increasingly written by agents, so the docs are built for them as well as for you.
        The Claude Code skill handles project setup, TSX authoring, preview and export.
        Install it once, then describe the presentation you want.
      </p>
    </div>
    <ul>
      <li>
        <a href="{base}/llms.txt"><code>/llms.txt</code></a>
        <span>An index of every docs page, for a model to pick from.</span>
      </li>
      <li>
        <a href="{base}/llms-full.txt"><code>/llms-full.txt</code></a>
        <span>The whole documentation in one file.</span>
      </li>
      <li>
        <a href="{base}/docs/getting-started.md"><code>any-docs-page.md</code></a>
        <span>Add .md to a docs URL to get the raw Markdown.</span>
      </li>
      <li>
        <a href="{currentProduct.repo}/blob/main/skill/SKILL.md"><code>skill/SKILL.md</code></a>
        <span>The installable Claude Code skill, with bundled TSX and core references.</span>
      </li>
    </ul>
  </div>
</section>

<section class="band" aria-labelledby="family-title">
  <div class="frame">
    <div class="section-head family-head">
      <h2 id="family-title">One kit, three file formats</h2>
      <p>
        Office Kit is a family of libraries built on the same rules: the ECMA-376 spec is the source
        of truth, output has to validate, and one ESM build has to run everywhere.
      </p>
    </div>
    <FamilyGrid product="pptx" />
  </div>
</section>

<style>
  /* Hero. */
  .hero-inner {
    padding: clamp(3rem, 6vw, 4.75rem) var(--gutter) clamp(2.75rem, 5vw, 3.75rem);
    text-align: center;
  }

  h1 {
    max-width: 21ch;
    margin: 0 auto;
    font-size: clamp(2.35rem, 6vw, 4.25rem);
    font-weight: 650;
    line-height: 1.02;
    letter-spacing: -0.04em;
  }

  .lede {
    max-width: 62ch;
    margin: 1.75rem auto 0;
    color: var(--ink-2);
    font-size: clamp(1.05rem, 1.6vw, 1.2rem);
    line-height: 1.55;
    text-wrap: pretty;
  }

  .cta {
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 0.6rem;
    margin-top: 2.25rem;
  }

  /* Stage: the product accent as a field, dotted like PowerPoint's canvas grid. */
  .stage {
    background-color: var(--accent);
    background-image: radial-gradient(circle, rgb(255 255 255 / 0.24) 1px, transparent 1.4px);
    background-size: 22px 22px;
  }

  .stage-inner {
    border-inline-color: rgb(255 255 255 / 0.22);
    padding: clamp(1.5rem, 4vw, 3.5rem) var(--gutter) 0;
  }

  .stage-grid {
    display: grid;
    grid-template-columns: minmax(0, 5fr) minmax(0, 7fr);
    gap: clamp(1rem, 2vw, 1.75rem);
    align-items: stretch;
  }

  figure {
    margin: 0;
  }

  .code-pane {
    display: flex;
    flex-direction: column;
    min-height: 0;
    border-radius: var(--radius);
    background: var(--night);
    box-shadow: var(--shadow-pop);
    overflow: hidden;
  }

  .code-pane figcaption,
  .template-code figcaption {
    flex: none;
    padding: 0.6rem 1rem;
    border-bottom: 1px solid var(--night-line);
    color: var(--night-ink-2);
    font-family: var(--mono);
    font-size: 0.8rem;
  }

  /* The pane's height comes from the slide beside it; the code scrolls inside. */
  .code-scroll {
    position: relative;
    flex: 1;
    min-height: 0;
    overflow: auto;
  }

  .code-scroll :global(pre) {
    position: absolute;
    inset: 0 auto auto 0;
    min-width: 100%;
    margin: 0;
    border: none;
    border-radius: 0;
    overflow: visible;
    font-size: 0.8rem;
  }

  .slide {
    aspect-ratio: 16 / 9;
    border-radius: 4px;
    background: #fff;
    box-shadow: var(--shadow-pop);
    overflow: hidden;
  }

  .slide :global(svg) {
    display: block;
    width: 100%;
    height: 100%;
  }

  .stage-foot {
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    gap: 1.5rem 3rem;
    padding: clamp(1.25rem, 3vw, 2rem) 0 clamp(1.5rem, 3vw, 2.5rem);
    color: var(--on-accent);
  }

  .stage-foot p {
    max-width: 62ch;
    margin: 0;
    font-size: 1rem;
    line-height: 1.55;
  }

  .stage-foot code {
    white-space: nowrap;
    background: rgb(255 255 255 / 0.16);
    border-color: rgb(255 255 255 / 0.3);
    color: inherit;
  }

  .stage-action {
    flex: none;
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    gap: 0.5rem;
    max-width: 22rem;
  }

  .stage-btn {
    background: #fff;
    border-color: #fff;
    color: #15171c;
  }

  .stage-btn:hover {
    background: #15171c;
    border-color: #15171c;
    color: #fff;
  }

  .stage-btn:disabled {
    cursor: progress;
    opacity: 0.8;
  }

  .stage-foot .stage-note {
    font-size: 0.88rem;
    text-align: right;
  }

  .stage :global(:focus-visible) {
    outline-color: #fff;
  }

  /* Three ways in. */
  .paths {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    list-style: none;
    margin: 0;
    padding: 0;
  }

  .paths li {
    display: flex;
    flex-direction: column;
    margin: 0;
    padding: 2.5rem var(--gutter) 2.25rem;
    border-right: 1px solid var(--line);
  }

  .paths li:last-child {
    border-right: none;
  }

  .path-code {
    align-self: flex-start;
    background: var(--accent-wash);
    border-color: transparent;
    color: var(--accent-ink);
    font-size: 0.82rem;
  }

  .paths h2 {
    margin: 1.1rem 0 0.6rem;
    font-size: 1.35rem;
  }

  .paths p {
    margin: 0 0 1.25rem;
    color: var(--ink-2);
    font-size: 0.97rem;
  }

  .paths a {
    margin-top: auto;
    font-weight: 550;
  }

  /* Template section. */
  .template-inner {
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
    gap: clamp(2rem, 5vw, 4.5rem);
    align-items: center;
    padding: clamp(3rem, 7vw, 5.5rem) var(--gutter);
  }

  .template-text h2 {
    margin: 0 0 1.25rem;
    font-size: clamp(1.7rem, 3.4vw, 2.5rem);
    letter-spacing: -0.03em;
    line-height: 1.08;
  }

  .template-text p {
    color: var(--ink-2);
    max-width: 52ch;
  }

  .template-code {
    border-radius: var(--radius);
    background: var(--night);
    border: 1px solid var(--night-line);
    overflow: hidden;
  }

  .template-code :global(pre) {
    margin: 0;
    border: none;
    border-radius: 0;
  }

  /* Proof: the one dark band on the page. */
  .proof {
    background: var(--night);
    color: var(--night-ink);
    border-bottom-color: var(--night-line);
  }

  .proof-inner {
    border-inline-color: var(--night-line);
    padding: clamp(3rem, 7vw, 5.5rem) var(--gutter);
  }

  .proof h2 {
    margin: 0;
    color: #fff;
    font-size: clamp(1.9rem, 4.2vw, 3.1rem);
    letter-spacing: -0.035em;
    line-height: 1.05;
  }

  .proof-lede {
    max-width: 58ch;
    margin: 1.1rem 0 2.75rem;
    color: var(--night-ink-2);
    font-size: 1.08rem;
  }

  .proof dl {
    margin: 0;
    border-top: 1px solid var(--night-line);
  }

  .proof-row {
    display: grid;
    grid-template-columns: minmax(0, 5fr) minmax(0, 6fr);
    gap: 0.5rem 3rem;
    padding: 1.5rem 0;
    border-bottom: 1px solid var(--night-line);
  }

  .proof dt {
    font-family: var(--display);
    font-size: 1.25rem;
    font-weight: 550;
    line-height: 1.25;
    letter-spacing: -0.015em;
    color: #fff;
    text-wrap: balance;
  }

  .proof dd {
    margin: 0;
    color: var(--night-ink-2);
    font-size: 0.98rem;
  }

  /* Section heads shared by the lower sections. */
  .section-head {
    max-width: 66ch;
    padding: clamp(3rem, 7vw, 5rem) var(--gutter) 0;
  }

  .section-head h2 {
    margin: 0 0 1rem;
    font-size: clamp(1.7rem, 3.4vw, 2.5rem);
    letter-spacing: -0.03em;
    line-height: 1.08;
  }

  .section-head p {
    margin: 0;
    color: var(--ink-2);
    font-size: 1.04rem;
  }

  /* Comparison. */
  .compare {
    padding-bottom: clamp(3rem, 6vw, 4.5rem);
  }

  .compare .table-scroll {
    margin-top: 2.25rem;
    padding: 0 var(--gutter);
  }

  .compare table {
    min-width: 640px;
    margin: 0;
    table-layout: fixed;
  }

  .compare thead th {
    width: 37%;
    padding-bottom: 0.8rem;
    font-family: var(--mono);
    font-size: 0.85rem;
    font-weight: 550;
    color: var(--ink);
  }

  .compare thead th:first-child {
    width: 26%;
  }

  .compare thead th:nth-child(2) {
    color: var(--accent-ink);
  }

  .compare tbody th {
    padding-right: 1.5rem;
    border-bottom-color: var(--line);
    font-weight: 600;
    color: var(--ink);
    white-space: normal;
  }

  .compare td {
    padding-right: 1.5rem;
    color: var(--ink-2);
  }

  /* A mark in front of the text carries yes / no, so colour is never the only cue. */
  .compare td[data-tone]::before {
    display: inline-block;
    width: 1.25em;
    font-weight: 700;
  }

  .compare td[data-tone='yes']::before {
    content: '✓';
    color: #168a4f;
  }

  .compare td[data-tone='no']::before {
    content: '✕';
    color: var(--ink-3);
  }

  .compare td[data-tone='part']::before {
    content: '◐';
    color: var(--ink-3);
  }

  .compare-advice {
    max-width: 70ch;
    margin: 2rem 0 0;
    padding: 0 var(--gutter);
    color: var(--ink);
    font-size: 1.02rem;
  }

  /* Capabilities. */
  .caps-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    margin: 2.5rem 0 0;
    border-top: 1px solid var(--line);
  }

  .caps-grid > div {
    padding: 1.5rem var(--gutter) 1.6rem;
    border-right: 1px solid var(--line);
    border-bottom: 1px solid var(--line);
  }

  .caps-grid > div:nth-child(4n) {
    border-right: none;
  }

  .caps-grid dt {
    font-family: var(--display);
    font-size: 1.12rem;
    font-weight: 600;
    letter-spacing: -0.01em;
  }

  .caps-grid dd {
    margin: 0.45rem 0 0;
    color: var(--ink-2);
    font-size: 0.93rem;
    line-height: 1.5;
  }

  .not-yet {
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(0, 3fr);
    gap: 0.75rem 2rem;
    padding: 1.75rem var(--gutter) 2.25rem;
    background: var(--wash);
  }

  .not-yet h3 {
    margin: 0;
    font-size: 1.12rem;
  }

  .not-yet ul {
    margin: 0;
    padding: 0;
    list-style: none;
    columns: 2;
    column-gap: 2.5rem;
    color: var(--ink-2);
    font-size: 0.93rem;
  }

  .not-yet li {
    margin: 0 0 0.45rem;
    break-inside: avoid;
  }

  /* Agents. */
  .agents-inner {
    display: grid;
    grid-template-columns: minmax(0, 5fr) minmax(0, 6fr);
    gap: clamp(2rem, 5vw, 4.5rem);
    padding: clamp(3rem, 7vw, 5rem) var(--gutter);
  }

  .agents h2 {
    margin: 0 0 1rem;
    font-size: clamp(1.7rem, 3.4vw, 2.5rem);
    letter-spacing: -0.03em;
    line-height: 1.08;
  }

  .agents p {
    color: var(--ink-2);
    max-width: 50ch;
  }

  .agents ul {
    list-style: none;
    margin: 0;
    padding: 0;
    border-top: 1px solid var(--line);
  }

  .agents li {
    display: grid;
    grid-template-columns: 11.5rem minmax(0, 1fr);
    gap: 0.25rem 1rem;
    align-items: baseline;
    margin: 0;
    padding: 0.95rem 0;
    border-bottom: 1px solid var(--line);
  }

  .agents li span {
    color: var(--ink-2);
    font-size: 0.95rem;
  }

  .family-head {
    padding-bottom: 2.5rem;
    border-bottom: 1px solid var(--line);
    max-width: none;
  }

  .family-head p {
    max-width: 66ch;
  }

  @media (max-width: 960px) {
    .stage-grid {
      grid-template-columns: 1fr;
    }

    /* Slide first on narrow screens: it is the payoff, the code is the proof. */
    .slide-pane {
      order: -1;
    }

    .code-scroll {
      height: 300px;
      flex: none;
    }

    .paths {
      grid-template-columns: 1fr;
    }

    .paths li {
      border-right: none;
      border-bottom: 1px solid var(--line);
      padding-block: 2rem;
    }

    .paths li:last-child {
      border-bottom: none;
    }

    .template-inner,
    .agents-inner {
      grid-template-columns: 1fr;
    }

    .caps-grid {
      grid-template-columns: repeat(2, 1fr);
    }

    .caps-grid > div:nth-child(2n) {
      border-right: none;
    }
  }

  @media (max-width: 720px) {
    .stage-foot {
      flex-direction: column;
    }

    .stage-action {
      align-items: flex-start;
      max-width: none;
    }

    .stage-foot .stage-note {
      text-align: left;
    }

    .proof-row {
      grid-template-columns: 1fr;
    }

    .not-yet {
      grid-template-columns: 1fr;
    }

    .not-yet ul {
      columns: 1;
    }
  }

  @media (max-width: 520px) {
    .cta .btn {
      flex: 1 1 100%;
    }

    .caps-grid {
      grid-template-columns: 1fr;
    }

    .caps-grid > div {
      border-right: none;
    }

    .agents li {
      grid-template-columns: 1fr;
    }
  }
</style>
