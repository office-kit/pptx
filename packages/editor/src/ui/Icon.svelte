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
  {:else if name === 'video'}
    <rect x="4" y="4" width="16" height="16" rx="2" /><path d="M8 4v16M16 4v16M4 8h4M16 8h4M4 16h4M16 16h4" />
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
  {:else if name === 'reset'}
    <path d="M4 8a8 8 0 1 1 0 9" /><polyline points="4,3 4,8 9,8" /><rect x="9" y="10" width="8" height="6" rx="1" />
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
  {:else if name === 'paste'}
    <rect x="4" y="4" width="12" height="16" rx="1.5" /><rect x="7.5" y="2.5" width="5" height="3" rx="1" /><rect x="11" y="10" width="9" height="11" rx="1" />
  {:else if name === 'cut'}
    <circle cx="7" cy="17" r="3" /><circle cx="17" cy="17" r="3" /><line x1="9" y1="15" x2="17" y2="3" /><line x1="15" y1="15" x2="7" y2="3" />
  {:else if name === 'copy'}
    <rect x="8" y="8" width="12" height="13" rx="1.5" /><path d="M5 16V4.5A1.5 1.5 0 0 1 6.5 3H15" />
  {:else if name === 'format-painter'}
    <rect x="4" y="3" width="13" height="6" rx="1" /><path d="M17 6h3v5h-8v3" /><rect x="10.5" y="14" width="3" height="7" rx="1" />
  {:else if name === 'new-slide'}
    <rect x="5" y="7" width="16" height="12" rx="1" /><line x1="5" y1="11" x2="21" y2="11" /><line x1="13" y1="11" x2="13" y2="19" /><line x1="4" y1="1.5" x2="4" y2="7.5" /><line x1="1" y1="4.5" x2="7" y2="4.5" />
  {:else if name === 'layout'}
    <rect x="3" y="5" width="18" height="14" rx="1" /><line x1="3" y1="9" x2="21" y2="9" /><line x1="12" y1="9" x2="12" y2="19" />
  {:else if name === 'section'}
    <rect x="4" y="3" width="16" height="4" rx="1" /><rect x="4" y="10" width="16" height="11" rx="1" />
  {:else if name === 'numbering'}
    <path d="M4 5h1.5v4" /><path d="M3.5 13.5a1.4 1.4 0 0 1 2.6.6c0 1-2.6 2-2.6 2.9h2.8" /><line x1="9" y1="7" x2="20" y2="7" /><line x1="9" y1="15" x2="20" y2="15" />
  {:else if name === 'indent-less'}
    <line x1="4" y1="5" x2="20" y2="5" /><line x1="11" y1="10" x2="20" y2="10" /><line x1="11" y1="14" x2="20" y2="14" /><line x1="4" y1="19" x2="20" y2="19" /><polyline points="8,9.5 5,12 8,14.5" />
  {:else if name === 'indent-more'}
    <line x1="4" y1="5" x2="20" y2="5" /><line x1="11" y1="10" x2="20" y2="10" /><line x1="11" y1="14" x2="20" y2="14" /><line x1="4" y1="19" x2="20" y2="19" /><polyline points="5,9.5 8,12 5,14.5" />
  {:else if name === 'arrange'}
    <rect x="3" y="3" width="9" height="9" /><rect x="8" y="8" width="9" height="9" fill="currentColor" opacity="0.35" /><rect x="12" y="12" width="9" height="9" />
  {:else if name === 'quick-styles'}
    <rect x="3" y="3" width="18" height="18" rx="2" /><path d="M8 17c2 0 3-1 3-3l6-7" /><circle cx="7.5" cy="17" r="1.5" />
  {:else if name === 'picture'}
    <rect x="3" y="5" width="18" height="14" rx="1" /><path d="M3 17l6-6 4 4 3-3 5 5" /><circle cx="16" cy="9" r="1.5" />
  {:else if name === 'shapes'}
    <rect x="3" y="3" width="11" height="11" /><circle cx="15" cy="15" r="6" />
  {:else if name === 'calendar'}
    <rect x="3" y="5" width="18" height="16" rx="1.5" /><line x1="3" y1="10" x2="21" y2="10" /><line x1="8" y1="3" x2="8" y2="7" /><line x1="16" y1="3" x2="16" y2="7" /><circle cx="15.5" cy="15.5" r="2.5" />
  {:else if name === 'slide-number'}
    <rect x="3" y="5" width="18" height="14" rx="1" /><path d="M10 9h4M12 9v6M10 15h4" /><path d="M9 11l1-2" />
  {:else if name === 'screenshot'}
    <rect x="3" y="4" width="14" height="10" rx="1" stroke-dasharray="2 2" /><rect x="10" y="13" width="9" height="7" rx="1" /><circle cx="14.5" cy="16.5" r="1.6" /><path d="M20 17v4M18 19h4" />
  {:else if name === 'cameo'}
    <rect x="3" y="4" width="18" height="14" rx="1" /><circle cx="15" cy="10" r="2.5" /><path d="M10.5 18c0-3 9-3 9 0" /><path d="M6 8h4M6 11h3" />
  {:else if name === 'icons'}
    <circle cx="12" cy="7" r="3" /><path d="M6 21c0-5 3-7 6-7s6 2 6 7" /><path d="M4 12l3-2M20 12l-3-2" />
  {:else if name === 'cube'}
    <path d="M12 3 20 7.5v9L12 21l-8-4.5v-9z" /><path d="M4 7.5 12 12l8-4.5M12 12v9" />
  {:else if name === 'font-color'}
    <path d="m7 15 5-11 5 11M8.8 11h6.4" /><path d="M4 19h16" stroke-width="3" />
  {:else if name === 'height'}
    <path d="M12 3v18M8 7l4-4 4 4M8 17l4 4 4-4" />
  {:else if name === 'width'}
    <path d="M3 12h18M7 8l-4 4 4 4M17 8l4 4-4 4" />
  {:else if name === 'format-pane'}
    <rect x="3" y="4" width="18" height="16" rx="1.5" /><path d="M14 4v16M16.5 8h2M16.5 11h2M16.5 14h2" />
  {:else if name === 'swap'}
    <path d="M4 8h13l-3-3M20 16H7l3 3" />
  {:else if name === 'share'}
    <path d="M14 4h6v6M20 4l-9 9M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5" />
  {:else if name === 'columns'}
    <path d="M3 6h7M3 10h7M3 14h7M3 18h7M14 6h7M14 10h7M14 14h7M14 18h5" />
  {:else if name === 'text-direction'}
    <path d="M6 4h8M10 4v9M18 7v13M15.5 17.5 18 20l2.5-2.5" />
  {:else if name === 'align-text'}
    <rect x="3" y="3" width="18" height="18" rx="2" /><path d="M7 10h10M7 14h10" />
  {:else if name === 'designer'}
    <rect x="3" y="7" width="18" height="13" rx="1.5" /><path d="m13 2-4 7h4l-2 6" />
  {:else if name === 'add-ins'}
    <rect x="3" y="3" width="8" height="8" rx="1" /><rect x="13" y="3" width="8" height="8" rx="1" /><rect x="3" y="13" width="8" height="8" rx="1" /><rect x="13" y="13" width="8" height="8" rx="1" />
  {:else if name === 'smartart'}
    <rect x="3" y="4" width="6" height="5" rx="1" /><rect x="15" y="4" width="6" height="5" rx="1" /><rect x="9" y="15" width="6" height="5" rx="1" /><path d="M6 9v3h12V9M12 12v3" />
  {:else if name === 'zoom-slide'}
    <rect x="3" y="5" width="12" height="9" rx="1" /><rect x="9" y="10" width="12" height="9" rx="1" /><path d="M7 9h4" />
  {:else if name === 'action'}
    <path d="M12 3l2.6 5.6 6 .7-4.5 4.1 1.2 6L12 16.5 6.7 19.4l1.2-6L3.4 9.3l6-.7z" />
  {:else if name === 'header-footer'}
    <path d="M6 3h9l4 4v14H6z" /><path d="M8 6h6M8 18h8" /><path d="M8 10h8M8 13h8" stroke-opacity="0.4" />
  {:else if name === 'text-art'}
    <path d="M5 20 11 4h2l6 16M8 14h8" /><path d="M3 21c6-2 12-2 18 0" />
  {:else if name === 'object'}
    <rect x="4" y="3" width="16" height="18" rx="1" /><rect x="7" y="7" width="10" height="7" /><path d="M7 17h6" />
  {:else if name === 'equation'}
    <path d="M5 5h12M8 5v14M14 5v11c0 3 3 3 4 2" />
  {:else if name === 'symbol'}
    <path d="M5 20h4v-2.5A7 7 0 1 1 15 17.5V20h4" />
  {:else if name === 'video'}
    <rect x="3" y="5" width="13" height="14" rx="1" /><path d="M16 10l5-3v10l-5-3" /><path d="M3 9h13M3 15h13" />
  {:else if name === 'audio'}
    <path d="M4 9h4l5-4v14l-5-4H4z" /><path d="M16 8.5a5 5 0 0 1 0 7M18.5 6a8.5 8.5 0 0 1 0 12" />
  {:else if name === 'rename'}
    <rect x="3" y="7" width="18" height="10" rx="1" /><path d="M7 10v4M14 5v14M12 5h4M12 19h4" />
  {:else if name === 'close-master'}
    <rect x="3" y="4" width="18" height="16" rx="1" /><path d="m9 9 6 6M15 9l-6 6" stroke="#c92a2a" />
  {:else if name === 'alt-text'}
    <rect x="3" y="4" width="14" height="11" rx="1" /><path d="m5 13 3-4 3 3 2-2 3 3" /><rect x="11" y="14" width="10" height="7" rx="1" /><path d="M13 17h6M13 19h4" />
  {:else if name === 'selection-pane'}
    <rect x="3" y="4" width="18" height="16" rx="1" /><path d="M14 4v16M16 8h3M16 11h3M16 14h3" /><path d="m6 8 4 9 1-3.5 3.5-1z" />
  {:else if name === 'align-objects'}
    <path d="M4 3v18" /><rect x="6" y="5" width="13" height="5" rx="1" /><rect x="6" y="14" width="8" height="5" rx="1" />
  {:else if name === 'remove-background'}
    <rect x="3" y="4" width="18" height="16" rx="1" stroke-dasharray="2 2" /><path d="M7 18c0-4 2-6 5-6s5 2 5 6" /><circle cx="12" cy="9" r="2.5" />
  {:else if name === 'corrections'}
    <circle cx="12" cy="12" r="4" /><path d="M12 2v3M12 19v3M2 12h3M19 12h3M5 5l2 2M17 17l2 2M5 19l2-2M17 7l2-2" />
  {:else if name === 'artistic-effects'}
    <rect x="3" y="4" width="18" height="16" rx="1" /><path d="M7 16c2-6 4-8 5-4s3 2 5-2" />
  {:else if name === 'transparency'}
    <rect x="3" y="4" width="18" height="16" rx="1" /><path d="M3 12h18M12 4v16" stroke-dasharray="2 2" />
  {:else if name === 'compress'}
    <rect x="5" y="5" width="14" height="14" rx="1" /><path d="M3 3l5 5M21 3l-5 5M3 21l5-5M21 21l-5-5" />
  {:else if name === 'gridlines'}
    <rect x="3" y="5" width="18" height="14" /><path d="M3 9.7h18M3 14.3h18M9 5v14M15 5v14" stroke-dasharray="1.5 1.5" />
  {:else if name === 'split'}
    <rect x="3" y="6" width="18" height="12" rx="1.5" /><path d="M12 6v12" /><polyline points="8,9.5 5.5,12 8,14.5" /><polyline points="16,9.5 18.5,12 16,14.5" />
  {:else if name === 'distribute-rows'}
    <rect x="4" y="4" width="16" height="16" /><path d="M4 9.3h16M4 14.6h16" />
  {:else if name === 'distribute-columns'}
    <rect x="4" y="4" width="16" height="16" /><path d="M9.3 4v16M14.6 4v16" />
  {:else if name === 'cell-margins'}
    <rect x="3" y="3" width="18" height="18" /><rect x="7" y="7" width="10" height="10" stroke-dasharray="2 2" />
  {:else if name === 'insert-row-above'}
    <rect x="3" y="11" width="18" height="9" /><path d="M3 15.5h18" /><path d="M12 9V3M9 6l3-3 3 3" />
  {:else if name === 'insert-row-below'}
    <rect x="3" y="4" width="18" height="9" /><path d="M3 8.5h18" /><path d="M12 15v6M9 18l3 3 3-3" />
  {:else if name === 'insert-column-left'}
    <rect x="11" y="3" width="9" height="18" /><path d="M15.5 3v18" /><path d="M9 12H3M6 9l-3 3 3 3" />
  {:else if name === 'insert-column-right'}
    <rect x="4" y="3" width="9" height="18" /><path d="M8.5 3v18" /><path d="M15 12h6M18 9l3 3-3 3" />
  {:else if name === 'select-table'}
    <path d="m5 3 5 15 2-6 6-2z" />
  {:else if name === 'pen'}
    <path d="M4 20l2-6L16 4l4 4-10 10z" /><path d="M14 6l4 4" />
  {:else if name === 'draw-table'}
    <rect x="3" y="6" width="13" height="13" /><path d="M3 12.5h13M9.5 6v13" /><path d="m14 10 6-6 1.5 1.5-6 6-2.5 1z" />
  {:else if name === 'eraser'}
    <path d="m4 16 9-9 7 7-5 5H8z" /><path d="M9 12l6 6M3 20h18" />
  {:else if name === 'align-top'}
    <path d="M4 4h16" /><rect x="7" y="7" width="4" height="12" /><rect x="13" y="7" width="4" height="7" />
  {:else if name === 'align-middle'}
    <path d="M4 12h16" /><rect x="7" y="5" width="4" height="14" /><rect x="13" y="8" width="4" height="8" />
  {:else if name === 'align-bottom'}
    <path d="M4 20h16" /><rect x="7" y="5" width="4" height="12" /><rect x="13" y="10" width="4" height="7" />
  {:else if name === 'animate-background'}
    <rect x="3" y="7" width="15" height="13" rx="1" /><path d="m5 18 4-5 3 3 2-2 3 4" /><path d="M17 3l1 2 2 1-2 1-1 2-1-2-2-1 2-1z" />
  {:else if name === 'picture-layout'}
    <rect x="3" y="4" width="8" height="7" /><rect x="13" y="4" width="8" height="7" /><path d="M3 15h18M3 19h12" />
  {:else}
    <circle cx="12" cy="12" r="2.5" fill="currentColor" stroke="none" />
  {/if}
</svg>
