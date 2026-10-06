// The ribbon's command area scrolls horizontally, and `overflow-x: auto` makes
// it clip vertically as well, so a dropdown positioned `absolute` inside it is
// cut off at the ribbon's bottom edge whatever its z-index. Ribbon dropdowns
// are `position: fixed` instead and placed by this action under the trigger
// button rendered right before them.
export function placeBelowTrigger(node: HTMLElement) {
  const trigger = node.previousElementSibling?.getBoundingClientRect();
  if (!trigger) return;
  node.style.left = `${Math.max(8, Math.min(trigger.left, innerWidth - node.offsetWidth - 8))}px`;
  node.style.top = `${trigger.bottom + 2}px`;
}
