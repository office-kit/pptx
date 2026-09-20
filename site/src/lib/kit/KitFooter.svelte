<script lang="ts">
  import { base } from '$app/paths';
  import KitMark from './KitMark.svelte';
  import { ORG_URL, currentProduct, products } from './products';

  type FooterLink = { path: string; label: string };

  type Props = {
    /** In-site links for the current product's column. */
    links: FooterLink[];
  };

  const { links }: Props = $props();
</script>

<footer class="kit-footer" data-pagefind-ignore>
  <div class="frame inner">
    <div class="about">
      <span class="wordmark">
        <KitMark front={currentProduct.id} size={24} />
        Office Kit
      </span>
      <p>
        TypeScript libraries for Office files. MIT licensed, no native modules, one ESM build for
        Node and the browser.
      </p>
    </div>

    <nav class="col" aria-label="{currentProduct.id} site">
      <h2>{currentProduct.id}</h2>
      <ul>
        {#each links as link (link.path)}
          <li><a href="{base}{link.path}">{link.label}</a></li>
        {/each}
        <li><a href="{currentProduct.repo}/blob/main/CHANGELOG.md">Changelog</a></li>
      </ul>
    </nav>

    <nav class="col" aria-label="Office Kit libraries">
      <h2>Libraries</h2>
      <ul>
        {#each products as p (p.id)}
          <li>
            <a href={p.id === currentProduct.id ? `${base}/` : p.href}>{p.pkg}</a>
          </li>
        {/each}
        <li><a href={ORG_URL}>GitHub organization</a></li>
      </ul>
    </nav>

    <nav class="col" aria-label="For AI agents">
      <h2>For AI agents</h2>
      <ul>
        <li><a href="{base}/llms.txt">llms.txt</a></li>
        <li><a href="{base}/llms-full.txt">llms-full.txt</a></li>
        <li><a href="{currentProduct.repo}/blob/main/skill/SKILL.md">Agent skill</a></li>
      </ul>
    </nav>
  </div>
</footer>

<style>
  .kit-footer {
    background: var(--wash);
  }

  .inner {
    display: grid;
    grid-template-columns: minmax(0, 1.6fr) repeat(3, minmax(0, 1fr));
    gap: 2rem;
    padding: 3rem var(--gutter) 3.5rem;
  }

  .wordmark {
    display: inline-flex;
    align-items: center;
    gap: 0.55rem;
    font-family: var(--display);
    font-weight: 650;
    font-size: 1.1rem;
    letter-spacing: -0.02em;
  }

  .about p {
    max-width: 34ch;
    margin: 0.9rem 0 0;
    color: var(--ink-2);
    font-size: 0.92rem;
  }

  h2 {
    margin: 0 0 0.8rem;
    font-family: var(--sans);
    font-size: 0.88rem;
    font-weight: 600;
    letter-spacing: 0;
    color: var(--ink-3);
  }

  ul {
    list-style: none;
    margin: 0;
    padding: 0;
  }

  li {
    margin: 0;
  }

  li a {
    display: inline-block;
    padding: 0.3rem 0;
    color: var(--ink);
    font-size: 0.94rem;
    overflow-wrap: anywhere;
  }

  @media (max-width: 860px) {
    .inner {
      grid-template-columns: repeat(3, minmax(0, 1fr));
      padding-block: 2.25rem 2.75rem;
    }

    .about {
      grid-column: 1 / -1;
    }
  }

  @media (max-width: 560px) {
    .inner {
      grid-template-columns: 1fr 1fr;
    }
  }
</style>
