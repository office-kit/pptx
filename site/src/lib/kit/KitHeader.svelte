<script lang="ts">
  import { base } from '$app/paths';
  import { page } from '$app/state';
  import { afterNavigate } from '$app/navigation';
  import Search from '$lib/components/Search.svelte';
  import KitMark from './KitMark.svelte';
  import ThemeToggle from './ThemeToggle.svelte';
  import { currentProduct, products } from './products';

  type NavLink = {
    path: string;
    label: string;
    /** Path prefix that marks this link current; defaults to `path`. */
    section?: string;
  };

  type Props = {
    /** In-site navigation for the current product. */
    links: NavLink[];
  };

  const { links }: Props = $props();

  let switcherOpen = $state(false);
  let menuOpen = $state(false);
  // eslint-disable-next-line prefer-const -- reassigned by `bind:this` in template
  let switcherEl = $state<HTMLDivElement | null>(null);

  // Sections nest (`/docs` contains `/docs/recipes`), so the longest matching
  // prefix wins and exactly one link is ever current.
  const activePath = $derived.by(() => {
    const here = page.url.pathname;
    let best: NavLink | undefined;
    for (const link of links) {
      const prefix = `${base}${link.section ?? link.path}`;
      if (!here.startsWith(prefix)) continue;
      if (!best || prefix.length > `${base}${best.section ?? best.path}`.length) best = link;
    }
    return best?.path;
  });

  function isActive(link: NavLink): boolean {
    return link.path === activePath;
  }

  function onDocumentClick(e: MouseEvent): void {
    if (switcherOpen && e.target instanceof Node && !switcherEl?.contains(e.target)) {
      switcherOpen = false;
    }
  }

  function onKeyDown(e: KeyboardEvent): void {
    if (e.key !== 'Escape') return;
    switcherOpen = false;
    menuOpen = false;
  }

  afterNavigate(() => {
    switcherOpen = false;
    menuOpen = false;
  });
</script>

<svelte:document onclick={onDocumentClick} onkeydown={onKeyDown} />

<header class="kit-header" data-pagefind-ignore>
  <div class="inner">
    <div class="brand">
      <a href="{base}/" class="brand-home" aria-label="{currentProduct.pkg} home">
        <KitMark front={currentProduct.id} />
        <span class="brand-name">Office Kit</span>
      </a>
      <span class="brand-slash" aria-hidden="true">/</span>
      <div class="switcher" bind:this={switcherEl}>
        <button
          type="button"
          class="switcher-btn"
          aria-haspopup="true"
          aria-expanded={switcherOpen}
          aria-controls="kit-product-menu"
          onclick={() => (switcherOpen = !switcherOpen)}
        >
          <span class="visually-hidden">Current library:</span>
          {currentProduct.id}
          <svg width="10" height="10" viewBox="0 0 10 10" aria-hidden="true">
            <path d="M1.5 3.5 5 7l3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.6" />
          </svg>
        </button>
        {#if switcherOpen}
          <ul id="kit-product-menu" class="switcher-menu">
            {#each products as p (p.id)}
              {@const current = p.id === currentProduct.id}
              <li>
                <a
                  href={current ? `${base}/` : p.href}
                  data-product={p.id}
                  aria-current={current ? 'true' : undefined}
                >
                  <span class="swatch" aria-hidden="true"></span>
                  <span class="product-text">
                    <span class="product-name">{p.id}</span>
                    <span class="product-app">{p.app} files</span>
                  </span>
                </a>
              </li>
            {/each}
          </ul>
        {/if}
      </div>
    </div>

    <nav class="nav" aria-label="Main">
      {#each links as link (link.path)}
        <a
          href="{base}{link.path}"
          class="nav-link"
          aria-current={isActive(link) ? 'page' : undefined}
        >
          {link.label}
        </a>
      {/each}
    </nav>

    <div class="tools">
      <Search />
      <ThemeToggle />
      <a class="icon-btn" href={currentProduct.repo} aria-label="{currentProduct.pkg} on GitHub">
        <svg width="20" height="20" viewBox="0 0 24 24" aria-hidden="true" fill="currentColor">
          <path
            d="M12 1.5a10.5 10.5 0 0 0-3.32 20.46c.53.1.72-.23.72-.5v-1.78c-2.92.64-3.54-1.4-3.54-1.4-.48-1.22-1.17-1.54-1.17-1.54-.95-.65.07-.64.07-.64 1.06.07 1.61 1.08 1.61 1.08.94 1.6 2.46 1.14 3.06.87.1-.68.37-1.14.67-1.4-2.33-.27-4.78-1.17-4.78-5.19 0-1.15.41-2.08 1.08-2.82-.11-.26-.47-1.33.1-2.78 0 0 .88-.28 2.89 1.08a10 10 0 0 1 5.26 0c2-1.36 2.88-1.08 2.88-1.08.58 1.45.22 2.52.11 2.78.67.74 1.08 1.67 1.08 2.82 0 4.03-2.46 4.92-4.8 5.18.38.33.72.97.72 1.96v2.9c0 .28.19.61.73.5A10.5 10.5 0 0 0 12 1.5Z"
          />
        </svg>
      </a>
      <button
        type="button"
        class="icon-btn menu-btn"
        aria-expanded={menuOpen}
        aria-controls="kit-mobile-menu"
        aria-label={menuOpen ? 'Close menu' : 'Open menu'}
        onclick={() => (menuOpen = !menuOpen)}
      >
        <svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true">
          {#if menuOpen}
            <path d="M5 5l10 10M15 5 5 15" stroke="currentColor" stroke-width="1.7" />
          {:else}
            <path d="M3 6h14M3 10h14M3 14h14" stroke="currentColor" stroke-width="1.7" />
          {/if}
        </svg>
      </button>
    </div>
  </div>

  {#if menuOpen}
    <div id="kit-mobile-menu" class="mobile-menu">
      <nav aria-label="Main">
        {#each links as link (link.path)}
          <a href="{base}{link.path}" aria-current={isActive(link) ? 'page' : undefined}>
            {link.label}
          </a>
        {/each}
      </nav>
      <p class="mobile-heading">Office Kit libraries</p>
      <ul class="mobile-products">
        {#each products as p (p.id)}
          {@const current = p.id === currentProduct.id}
          <li>
            <a
              href={current ? `${base}/` : p.href}
              data-product={p.id}
              aria-current={current ? 'true' : undefined}
            >
              <span class="swatch" aria-hidden="true"></span>
              <span class="product-name">{p.id}</span>
              <span class="product-app">{p.app} files</span>
            </a>
          </li>
        {/each}
      </ul>
    </div>
  {/if}
</header>

<style>
  .kit-header {
    position: sticky;
    top: 0;
    z-index: 30;
    background: color-mix(in srgb, var(--paper) 88%, transparent);
    backdrop-filter: blur(12px) saturate(1.4);
    -webkit-backdrop-filter: blur(12px) saturate(1.4);
    border-bottom: 1px solid var(--line);
  }

  .inner {
    height: var(--header-h);
    max-width: var(--frame);
    margin: 0 auto;
    padding: 0 var(--gutter);
    display: flex;
    align-items: center;
    gap: 1.75rem;
  }

  .brand {
    display: flex;
    align-items: center;
    gap: 0.45rem;
    flex: none;
  }

  .brand-home {
    display: inline-flex;
    align-items: center;
    gap: 0.55rem;
    color: var(--ink);
  }

  .brand-home:hover {
    text-decoration: none;
  }

  .brand-name {
    font-family: var(--display);
    font-weight: 650;
    font-size: 1.12rem;
    letter-spacing: -0.02em;
    white-space: nowrap;
  }

  .brand-slash {
    color: var(--line-strong);
    font-size: 1.25rem;
    font-weight: 300;
  }

  .switcher {
    position: relative;
  }

  .switcher-btn {
    display: inline-flex;
    align-items: center;
    gap: 0.4rem;
    height: 32px;
    padding: 0 0.55rem 0 0.6rem;
    border: 1px solid transparent;
    border-radius: var(--radius-sm);
    background: var(--accent-wash);
    color: var(--accent-ink);
    font-family: var(--mono);
    font-size: 0.9rem;
    font-weight: 550;
    cursor: pointer;
  }

  .switcher-btn:hover,
  .switcher-btn[aria-expanded='true'] {
    border-color: var(--accent);
  }

  .switcher-menu {
    position: absolute;
    top: calc(100% + 8px);
    left: 0;
    width: 232px;
    margin: 0;
    padding: 0.35rem;
    list-style: none;
    background: var(--paper);
    border: 1px solid var(--line-strong);
    border-radius: var(--radius);
    box-shadow: var(--shadow-pop);
  }

  .switcher-menu li {
    margin: 0;
  }

  .switcher-menu a,
  .mobile-products a {
    display: flex;
    align-items: center;
    gap: 0.7rem;
    padding: 0.55rem 0.6rem;
    border-radius: var(--radius-sm);
    color: var(--ink);
  }

  .switcher-menu a:hover,
  .mobile-products a:hover {
    background: var(--wash);
    text-decoration: none;
  }

  .switcher-menu a[aria-current],
  .mobile-products a[aria-current] {
    background: var(--accent-wash);
  }

  .swatch {
    flex: none;
    width: 12px;
    height: 12px;
    border-radius: 3px;
    background: var(--accent);
  }

  .product-text {
    display: flex;
    flex-direction: column;
    line-height: 1.25;
  }

  .product-name {
    font-family: var(--mono);
    font-weight: 550;
    font-size: 0.92rem;
  }

  .product-app {
    font-size: 0.82rem;
    color: var(--ink-2);
  }

  .nav {
    display: flex;
    align-items: center;
    gap: 0.15rem;
    min-width: 0;
  }

  .nav-link {
    padding: 0.4rem 0.7rem;
    border-radius: var(--radius-sm);
    color: var(--ink-2);
    font-size: 0.95rem;
    font-weight: 500;
    white-space: nowrap;
  }

  .nav-link:hover {
    color: var(--ink);
    background: var(--wash);
    text-decoration: none;
  }

  .nav-link[aria-current='page'] {
    color: var(--ink);
    font-weight: 600;
  }

  .tools {
    display: flex;
    align-items: center;
    gap: 0.25rem;
    margin-left: auto;
  }

  .icon-btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 40px;
    height: 40px;
    padding: 0;
    border: none;
    border-radius: var(--radius-sm);
    background: transparent;
    color: var(--ink-2);
    cursor: pointer;
  }

  .icon-btn:hover {
    color: var(--ink);
    background: var(--wash);
  }

  .menu-btn {
    display: none;
  }

  .mobile-menu {
    display: none;
  }

  @media (max-width: 900px) {
    .nav {
      display: none;
    }

    .menu-btn {
      display: inline-flex;
    }

    .mobile-menu {
      display: block;
      position: absolute;
      inset: 100% 0 auto;
      max-height: calc(100dvh - var(--header-h));
      overflow-y: auto;
      padding: 0.75rem var(--gutter) 1.5rem;
      background: var(--paper);
      border-bottom: 1px solid var(--line);
      box-shadow: var(--shadow-pop);
    }

    .mobile-menu nav {
      display: flex;
      flex-direction: column;
    }

    .mobile-menu nav a {
      padding: 0.8rem 0;
      border-bottom: 1px solid var(--line);
      color: var(--ink);
      font-size: 1.05rem;
      font-weight: 500;
    }

    .mobile-menu nav a[aria-current='page'] {
      color: var(--accent-ink);
    }

    .mobile-heading {
      margin: 1.5rem 0 0.4rem;
      font-size: 0.85rem;
      color: var(--ink-3);
    }

    .mobile-products {
      list-style: none;
      margin: 0;
      padding: 0;
    }

    .mobile-products li {
      margin: 0;
    }

    .mobile-products .product-app {
      margin-left: auto;
    }
  }

  @media (max-width: 520px) {
    .inner {
      gap: 0.75rem;
      padding: 0 1rem;
    }

    /* The mark plus the product chip still say "Office Kit / pptx". */
    .brand-name,
    .brand-slash {
      display: none;
    }
  }
</style>
