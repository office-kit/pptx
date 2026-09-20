<script lang="ts">
  import type { ProductId } from './products';

  type Props = {
    /** The product drawn on top of the stack. */
    front?: ProductId;
    size?: number;
  };

  const { front = 'pptx', size = 26 }: Props = $props();

  const FILL: Record<ProductId, string> = {
    pptx: '#e5481f',
    xlsx: '#168a4f',
    docx: '#2b63d9',
  };

  // Back-to-front paint order, ending on the product this site documents.
  const order = $derived<ProductId[]>([
    ...(['docx', 'xlsx', 'pptx'] as const).filter((id) => id !== front),
    front,
  ]);
</script>

<!-- Three stacked files: the kit. Decorative; the wordmark next to it carries the name. -->
<svg width={size} height={size} viewBox="0 0 26 26" aria-hidden="true" focusable="false">
  {#each order as id, i (id)}
    <rect
      x={1 + i * 4.5}
      y={10 - i * 4.5}
      width="15"
      height="15"
      rx="3.5"
      fill={FILL[id]}
      stroke="var(--paper)"
      stroke-width="1.5"
    />
  {/each}
</svg>
