<script lang="ts">
  // A small stroke-icon set (24×24, currentColor) so the ribbon and menus read
  // as a modern editor rather than an emoji collage. Icons are drawn from simple
  // geometric primitives for reliability. Unknown names render a neutral dot.
  interface Props {
    name: string;
    size?: number;
  }
  let { name, size = 20 }: Props = $props();
</script>

<svg
  width={size}
  height={size}
  viewBox="0 0 24 24"
  fill="none"
  stroke="currentColor"
  stroke-width="1.7"
  stroke-linecap="round"
  stroke-linejoin="round"
  aria-hidden="true"
>
  {#if name === 'slide-blank'}
    <rect x="3" y="5" width="18" height="14" rx="2" />
  {:else if name === 'slide-title'}
    <rect x="3" y="5" width="18" height="14" rx="2" /><line x1="7" y1="10" x2="17" y2="10" /><line x1="7" y1="14" x2="13" y2="14" />
  {:else if name === 'slide-content'}
    <rect x="3" y="5" width="18" height="14" rx="2" /><line x1="7" y1="9" x2="17" y2="9" /><line x1="7" y1="12.5" x2="17" y2="12.5" /><line x1="7" y1="16" x2="14" y2="16" />
  {:else if name === 'duplicate'}
    <rect x="9" y="9" width="11" height="11" rx="2" /><path d="M5 15V5a2 2 0 0 1 2-2h8" />
  {:else if name === 'trash'}
    <path d="M4 7h16" /><path d="M9 7V5a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2" /><path d="M6 7l1 12a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-12" />
  {:else if name === 'text-format'}
    <path d="M6 20l5-14h2l5 14" /><line x1="8.5" y1="14" x2="15.5" y2="14" />
  {:else if name === 'text-run'}
    <path d="M7 18l3.5-10h1L15 18" /><line x1="8.6" y1="14" x2="13.4" y2="14" /><line x1="17" y1="8" x2="17" y2="18" />
  {:else if name === 'align'}
    <line x1="4" y1="6" x2="20" y2="6" /><line x1="7" y1="11" x2="17" y2="11" /><line x1="4" y1="16" x2="20" y2="16" />
  {:else if name === 'bullets'}
    <circle cx="5" cy="7" r="1.2" /><circle cx="5" cy="12" r="1.2" /><circle cx="5" cy="17" r="1.2" /><line x1="9" y1="7" x2="20" y2="7" /><line x1="9" y1="12" x2="20" y2="12" /><line x1="9" y1="17" x2="20" y2="17" />
  {:else if name === 'indent'}
    <line x1="4" y1="6" x2="20" y2="6" /><line x1="10" y1="11" x2="20" y2="11" /><line x1="10" y1="16" x2="20" y2="16" /><polyline points="4,9 7,11.5 4,14" />
  {:else if name === 'line-spacing'}
    <line x1="10" y1="7" x2="20" y2="7" /><line x1="10" y1="12" x2="20" y2="12" /><line x1="10" y1="17" x2="20" y2="17" /><polyline points="5,9 5,15" /><polyline points="3,10.5 5,8 7,10.5" /><polyline points="3,13.5 5,16 7,13.5" />
  {:else if name === 'space'}
    <line x1="4" y1="8" x2="20" y2="8" /><line x1="4" y1="16" x2="20" y2="16" /><polyline points="12,10.5 12,13.5" />
  {:else if name === 'shape'}
    <rect x="4" y="4" width="16" height="16" rx="2" />
  {:else if name === 'textbox'}
    <rect x="3" y="6" width="18" height="12" rx="1" stroke-dasharray="3 2" /><line x1="12" y1="9" x2="12" y2="15" /><line x1="9.5" y1="9" x2="14.5" y2="9" />
  {:else if name === 'fill'}
    <path d="M5 11l6-6 7 7-6 6a2 2 0 0 1-3 0l-4-4a2 2 0 0 1 0-3z" /><path d="M18 15c1 1.5 2 2.5 2 3.5a2 2 0 1 1-4 0c0-1 1-2 2-3.5z" />
  {:else if name === 'outline'}
    <rect x="4" y="4" width="16" height="16" rx="2" stroke-width="2.6" />
  {:else if name === 'no-fill'}
    <rect x="4" y="4" width="16" height="16" rx="2" /><line x1="6" y1="18" x2="18" y2="6" stroke="#c0392b" />
  {:else if name === 'shadow'}
    <rect x="5" y="5" width="11" height="11" rx="1.5" /><path d="M9 20h9a2 2 0 0 0 2-2V9" opacity="0.5" />
  {:else if name === 'glow'}
    <circle cx="12" cy="12" r="4" /><line x1="12" y1="3" x2="12" y2="5.5" /><line x1="12" y1="18.5" x2="12" y2="21" /><line x1="3" y1="12" x2="5.5" y2="12" /><line x1="18.5" y1="12" x2="21" y2="12" /><line x1="5.6" y1="5.6" x2="7.4" y2="7.4" /><line x1="16.6" y1="16.6" x2="18.4" y2="18.4" /><line x1="18.4" y1="5.6" x2="16.6" y2="7.4" /><line x1="7.4" y1="16.6" x2="5.6" y2="18.4" />
  {:else if name === 'gradient'}
    <rect x="4" y="4" width="16" height="16" rx="2" /><defs><linearGradient id="g" x1="0" x2="1"><stop offset="0" stop-color="currentColor" stop-opacity="0.1" /><stop offset="1" stop-color="currentColor" stop-opacity="0.6" /></linearGradient></defs><rect x="4" y="4" width="16" height="16" rx="2" fill="url(#g)" stroke="none" />
  {:else if name === 'pattern'}
    <rect x="4" y="4" width="16" height="16" rx="2" /><line x1="4" y1="12" x2="12" y2="4" /><line x1="8" y1="20" x2="20" y2="8" /><line x1="14" y1="20" x2="20" y2="14" />
  {:else if name === 'front'}
    <rect x="7" y="7" width="12" height="12" rx="1.5" opacity="0.45" /><rect x="4" y="4" width="12" height="12" rx="1.5" />
  {:else if name === 'back'}
    <rect x="4" y="4" width="12" height="12" rx="1.5" opacity="0.45" /><rect x="7" y="7" width="12" height="12" rx="1.5" />
  {:else if name === 'forward'}
    <rect x="4" y="4" width="16" height="16" rx="2" /><polyline points="9,13 12,10 15,13" />
  {:else if name === 'backward'}
    <rect x="4" y="4" width="16" height="16" rx="2" /><polyline points="9,11 12,14 15,11" />
  {:else if name === 'group'}
    <rect x="4" y="4" width="7" height="7" rx="1" /><rect x="13" y="13" width="7" height="7" rx="1" /><path d="M11 7h4a2 2 0 0 1 2 2v4" stroke-dasharray="2 2" />
  {:else if name === 'ungroup'}
    <rect x="3" y="6" width="7" height="7" rx="1" /><rect x="14" y="11" width="7" height="7" rx="1" />
  {:else if name === 'replace'}
    <polyline points="4,8 8,4 8,12" /><path d="M8 8h8a4 4 0 0 1 0 8" /><polyline points="20,16 16,20 16,12" />
  {:else if name === 'table'}
    <rect x="3" y="4" width="18" height="16" rx="1.5" /><line x1="3" y1="10" x2="21" y2="10" /><line x1="3" y1="15" x2="21" y2="15" /><line x1="9" y1="4" x2="9" y2="20" /><line x1="15" y1="4" x2="15" y2="20" />
  {:else if name === 'image'}
    <rect x="3" y="4" width="18" height="16" rx="2" /><circle cx="8.5" cy="9.5" r="1.6" /><path d="M4 18l5-5 4 4 3-3 4 4" />
  {:else if name === 'chart'}
    <line x1="4" y1="20" x2="20" y2="20" /><rect x="6" y="11" width="3" height="7" /><rect x="11" y="7" width="3" height="11" /><rect x="16" y="14" width="3" height="4" />
  {:else if name === 'line'}
    <line x1="4" y1="20" x2="20" y2="4" />
  {:else if name === 'link'}
    <path d="M9 15l6-6" /><path d="M11 6l1-1a4 4 0 0 1 6 6l-1 1" /><path d="M13 18l-1 1a4 4 0 0 1-6-6l1-1" />
  {:else if name === 'comment'}
    <path d="M20 15a2 2 0 0 1-2 2H8l-4 4V6a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2z" />
  {:else if name === 'background'}
    <rect x="3" y="4" width="18" height="16" rx="2" /><path d="M3 15l5-5 4 4 3-3 6 6" opacity="0.6" />
  {:else if name === 'theme'}
    <circle cx="12" cy="12" r="8" /><circle cx="9" cy="9" r="1.2" fill="currentColor" stroke="none" /><circle cx="15" cy="9" r="1.2" fill="currentColor" stroke="none" /><circle cx="9" cy="15" r="1.2" fill="currentColor" stroke="none" /><circle cx="15" cy="15" r="1.2" fill="currentColor" stroke="none" />
  {:else if name === 'font'}
    <path d="M5 18l4-12h2l4 12" /><line x1="7" y1="13" x2="13" y2="13" /><path d="M16 18h3" />
  {:else if name === 'transition'}
    <rect x="3" y="6" width="7" height="12" rx="1" /><rect x="14" y="6" width="7" height="12" rx="1" opacity="0.5" /><polyline points="10.5,10 13.5,12 10.5,14" />
  {:else if name === 'animation'}
    <path d="M12 3l1.8 4.6L18.5 9l-4.7 1.4L12 15l-1.8-4.6L5.5 9l4.7-1.4z" /><circle cx="18" cy="17" r="1.4" /><circle cx="6" cy="16" r="1" />
  {:else if name === 'crop'}
    <path d="M6 2v14a2 2 0 0 0 2 2h14" /><path d="M2 6h14a2 2 0 0 1 2 2v14" />
  {:else if name === 'rotate'}
    <path d="M20 11a8 8 0 1 0-2.3 6.3" /><polyline points="20,5 20,11 14,11" />
  {:else if name === 'flip'}
    <line x1="12" y1="3" x2="12" y2="21" stroke-dasharray="2 2" /><path d="M9 7l-4 5 4 5z" /><path d="M15 7l4 5-4 5z" opacity="0.5" />
  {:else if name === 'dash'}
    <line x1="4" y1="12" x2="20" y2="12" stroke-dasharray="4 3" />
  {:else if name === 'arrow'}
    <line x1="4" y1="12" x2="18" y2="12" /><polyline points="13,7 19,12 13,17" />
  {:else if name === 'resize'}
    <path d="M4 10V4h6" /><path d="M20 14v6h-6" /><line x1="4" y1="4" x2="10" y2="10" /><line x1="20" y1="20" x2="14" y2="14" />
  {:else if name === 'cells-row'}
    <rect x="3" y="5" width="18" height="14" rx="1.5" /><line x1="3" y1="12" x2="21" y2="12" /><polyline points="12,8 12,9.5" /><polyline points="12,14.5 12,16" />
  {:else if name === 'cells-col'}
    <rect x="3" y="5" width="18" height="14" rx="1.5" /><line x1="12" y1="5" x2="12" y2="19" /><polyline points="8,12 9.5,12" /><polyline points="14.5,12 16,12" />
  {:else if name === 'merge'}
    <rect x="3" y="6" width="18" height="12" rx="1.5" /><path d="M9 12h6" /><polyline points="7,9 4,12 7,15" opacity="0" /><polyline points="10,9.5 12,12 10,14.5" /><polyline points="14,9.5 12,12 14,14.5" />
  {:else if name === 'border'}
    <rect x="4" y="4" width="16" height="16" /><line x1="4" y1="12" x2="20" y2="12" stroke-dasharray="2 2" /><line x1="12" y1="4" x2="12" y2="20" stroke-dasharray="2 2" />
  {:else}
    <circle cx="12" cy="12" r="2.5" fill="currentColor" stroke="none" />
  {/if}
</svg>
