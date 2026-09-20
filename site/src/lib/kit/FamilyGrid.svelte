<script lang="ts">
  import { base } from '$app/paths';
  import { currentProduct, products } from './products';
</script>

<!-- The three libraries side by side. Each cell takes its own product accent,
     so the grid doubles as the key to the colour coding used across the sites. -->
<ul class="family">
  {#each products as p (p.id)}
    {@const current = p.id === currentProduct.id}
    <li data-product={p.id}>
      <a href={current ? `${base}/` : p.href} aria-current={current ? 'true' : undefined}>
        <span class="ext">.{p.id}</span>
        <span class="pkg">{p.pkg}</span>
        <span class="summary">{p.app} files. {p.summary}</span>
        <span class="state">{current ? 'You are here' : `Open the ${p.id} site`}</span>
      </a>
    </li>
  {/each}
</ul>

<style>
  .family {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    list-style: none;
    margin: 0;
    padding: 0;
  }

  li {
    margin: 0;
    border-right: 1px solid var(--line);
  }

  li:last-child {
    border-right: none;
  }

  a {
    position: relative;
    display: flex;
    flex-direction: column;
    gap: 0.35rem;
    height: 100%;
    padding: 2rem var(--gutter) 1.75rem;
    color: var(--ink);
  }

  a::before {
    content: '';
    position: absolute;
    inset: 0 0 auto;
    height: 3px;
    background: var(--accent);
    transform: scaleX(0);
    transform-origin: left;
    transition: transform 200ms ease;
  }

  a:hover {
    text-decoration: none;
    background: var(--wash);
  }

  a:hover::before,
  a[aria-current]::before {
    transform: scaleX(1);
  }

  .ext {
    font-family: var(--display);
    font-size: clamp(2rem, 4vw, 2.75rem);
    font-weight: 650;
    letter-spacing: -0.035em;
    line-height: 1;
    color: var(--accent-ink);
  }

  .pkg {
    margin-top: 0.6rem;
    font-family: var(--mono);
    font-size: 0.88rem;
    font-weight: 500;
  }

  .summary {
    color: var(--ink-2);
    font-size: 0.95rem;
    line-height: 1.5;
  }

  .state {
    margin-top: auto;
    padding-top: 1.1rem;
    font-size: 0.88rem;
    font-weight: 550;
    color: var(--accent-ink);
  }

  a[aria-current] .state {
    color: var(--ink-3);
    font-weight: 450;
  }

  @media (max-width: 760px) {
    .family {
      grid-template-columns: 1fr;
    }

    li {
      border-right: none;
      border-bottom: 1px solid var(--line);
    }

    li:last-child {
      border-bottom: none;
    }

    a {
      padding: 1.5rem var(--gutter);
    }
  }
</style>
