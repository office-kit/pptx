import { transitionEffects, transitionTile } from './transition-gallery.ts';
import {
  editingHome,
  editingTabs,
  editingPanels,
  slideSizePanel,
  contextualTabs,
  contextualPanels,
} from './editor-ui.ts';
// Office-style chrome around the existing preview and agent editing workflow.
const paths = {
  file: '<path d="M6 3h9l4 4v14H6zM14 3v5h5M9 12h7M9 16h7"/>',
  play: '<rect x="3" y="4" width="18" height="13" rx="1"/><path d="m10 8 5 3-5 3zM12 17v4M8 21h8"/>',
  first: '<path d="M4 5v14M8 5l12 7-12 7z"/>',
  previous: '<path d="m15 5-7 7 7 7"/>',
  next: '<path d="m9 5 7 7-7 7"/>',
  normal:
    '<rect x="3" y="4" width="18" height="16" rx="1"/><path d="M8 4v16M3 8h5M3 12h5M3 16h5"/>',
  fit: '<rect x="6" y="7" width="12" height="10"/><path d="M3 9V4h5M16 4h5v5M21 15v5h-5M8 20H3v-5"/>',
  zoom: '<circle cx="10" cy="10" r="6"/><path d="m15 15 6 6"/>',
  agent: '<path d="M3 4h18v13H9l-5 4v-4H3zM7 8h10M7 12h7"/>',
} as const;
export function officeIcon(name: keyof typeof paths) {
  return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round" stroke-linecap="round" aria-hidden="true">${paths[name]}</svg>`;
}
const command = (id: string, label: string, icon: keyof typeof paths, extra = '') =>
  `<button id="${id}" class="ribbon-command" ${extra}>${officeIcon(icon)}<span>${label}</span></button>`;
export const officeHeader = `<header class="office-header">
<div class="application-menubar" role="menubar" aria-label="Application"><button id="file-menu" role="menuitem" data-edit="file-menu" aria-haspopup="menu" aria-expanded="false" aria-controls="office-menu">File</button><button id="edit-menu" role="menuitem" data-edit="edit-menu" aria-haspopup="menu" aria-expanded="false" aria-controls="office-menu" tabindex="-1">Edit</button><button id="view-menu" role="menuitem" data-edit="view-menu" aria-haspopup="menu" aria-expanded="false" aria-controls="office-menu" tabindex="-1">View</button><button id="format-menu" role="menuitem" data-edit="format-menu" aria-haspopup="menu" aria-expanded="false" aria-controls="office-menu" tabindex="-1">Format</button><button id="arrange-menu" role="menuitem" data-edit="arrange-menu" aria-haspopup="menu" aria-expanded="false" aria-controls="office-menu" tabindex="-1">Arrange</button></div>
<div class="titlebar"><span class="app-mark">P</span><a class="quick-save" href="/deck.pptx" title="Download PowerPoint file" aria-label="Download PPTX">${officeIcon('file')}</a><button class="quick-edit" data-edit="undo" title="Undo" aria-label="Undo">↶</button><button class="quick-edit" data-edit="redo" title="Redo" aria-label="Redo">↷</button><span class="document-title">Presentation <span>— Office Kit</span></span><div class="presentation-search" role="search" aria-label="Search in Presentation"><button id="find-menu" aria-label="Search options" aria-haspopup="menu" aria-expanded="false" aria-controls="office-menu">⌕⌄</button><input id="find-text" type="search" placeholder="Search in Presentation" aria-label="Search in Presentation"><button id="find-previous" aria-label="Find Previous" title="Find Previous">‹</button><button id="find-next" aria-label="Find Next" title="Find Next">›</button><span id="find-status" role="status" aria-live="polite"></span></div><span id="status" role="status">Building…</span></div>
<div class="tabbar"><div role="tablist" aria-label="Ribbon tabs"><button role="tab" id="tab-home" aria-controls="ribbon-home" aria-selected="true">Home</button>${editingTabs}<button role="tab" id="tab-transitions" aria-controls="ribbon-transitions" aria-selected="false" tabindex="-1">Transitions</button><button role="tab" id="tab-animations" aria-controls="ribbon-animations" aria-selected="false" tabindex="-1">Animations</button><button role="tab" id="tab-show" aria-controls="ribbon-show" aria-selected="false" tabindex="-1">Slide Show</button><button role="tab" id="tab-view" aria-controls="ribbon-view" aria-selected="false" tabindex="-1">View</button>${contextualTabs}</div><button id="toggle-chat" aria-expanded="true" aria-controls="chat">${officeIcon('agent')} Agents</button><a class="download" href="/deck.pptx">Export PPTX</a></div>
<div class="ribbon" id="ribbon-home" role="tabpanel" aria-labelledby="tab-home">${editingHome}<div hidden class="ribbon-group">${command('home-export', 'Export<br>PowerPoint', 'file', 'data-export')}<span class="group-label">File</span></div><div class="ribbon-group"><div class="ribbon-stack"><button id="prev" aria-label="Previous slide">${officeIcon('previous')} Previous Slide</button><button id="next" aria-label="Next slide">${officeIcon('next')} Next Slide</button></div><span class="group-label">Slides</span></div><div class="ribbon-group">${command('present', 'Present', 'play', 'disabled')}<span class="group-label">Slide Show</span></div><div class="ribbon-group">${command('home-agents', 'Agents', 'agent')}<span class="group-label">Editing</span></div></div>
${editingPanels}${slideSizePanel}${contextualPanels}
<div class="ribbon" id="ribbon-animations" role="tabpanel" aria-labelledby="tab-animations" hidden><div class="ribbon-group">${command('animation-preview', 'Preview', 'play', 'data-edit="animation-preview"')}<span class="group-label">Preview</span></div>${(
  ['Entrance', 'Exit'] as const
)
  .map(
    (category) =>
      `<div class="ribbon-group"><div class="transition-gallery" aria-label="${category} Effects">${(category ===
      'Entrance'
        ? [
            ['appear', 'Appear'],
            ['fadeIn', 'Fade'],
          ]
        : [
            ['disappear', 'Disappear'],
            ['fadeOut', 'Fade'],
          ]
      )
        .map(
          ([effect, label]) =>
            `<button class="ribbon-command" data-edit="animation-add" data-effect="${effect}" aria-label="${category} ${label}"><svg viewBox="0 0 24 24" width="30" height="30" aria-hidden="true"><path fill="${category === 'Entrance' ? '#6baf45' : '#cf5d62'}" stroke="${category === 'Entrance' ? '#488533' : '#a94246'}" d="m12 2 3 6 7 1-5 5 1 7-6-3-6 3 1-7-5-5 7-1z"/></svg><span>${label}</span></button>`,
        )
        .join('')}</div><span class="group-label">${category} Effects</span></div>`,
  )
  .join(
    '',
  )}<div class="ribbon-group">${command('animation-pane-toggle', 'Animation<br>Pane', 'normal', 'data-edit="animation-pane" aria-controls="animation-pane" aria-expanded="false"')}<span class="group-label">Advanced Animation</span></div><div class="ribbon-group"><div class="ribbon-stack"><label>Start: <select id="animation-start" aria-label="Animation start" disabled><option value="clickEffect">On Click</option><option value="withEffect">With Previous</option><option value="afterEffect">After Previous</option></select></label><label>Duration: <input id="animation-duration" aria-label="Animation duration" type="number" min="0" max="4294967.295" step="0.01" style="width:76px" disabled></label><label>Delay: <input id="animation-delay" aria-label="Animation delay" type="number" min="0" max="4294967.295" step="0.01" style="width:76px" disabled></label></div><span class="group-label">Timing</span></div></div>
<div class="ribbon" id="ribbon-transitions" role="tabpanel" aria-labelledby="tab-transitions" hidden><div class="ribbon-group">${command('transition-preview', 'Preview', 'play', 'data-edit="transition-preview"')}<span class="group-label">Preview</span></div><div class="ribbon-group"><div class="transition-gallery" aria-label="Transition effects">${transitionEffects
  .slice(0, 6)
  .map(([effect, label]) => transitionTile(effect, label))
  .join(
    '',
  )}</div><button data-edit="transition-more" aria-label="More transition effects" aria-haspopup="menu">▾</button>${command('transition-options', 'Effect<br>Options', 'next', 'data-edit="transition-options" aria-haspopup="menu"')}<span class="group-label">Transition to This Slide</span></div><div class="ribbon-group"><div class="ribbon-stack"><label>Sound: <select id="transition-sound" aria-label="Transition sound"><option value="none">[No Sound]</option><option value="stop">[Stop Previous Sound]</option><option value="other">Other Sound...</option></select></label><input id="transition-sound-file" type="file" accept=".wav,audio/wav,audio/x-wav" hidden><label class="ribbon-check"><input id="transition-sound-loop" type="checkbox"> Loop Until Next Sound</label><label>Duration: <input id="transition-duration" aria-label="Transition duration" type="number" min="0" max="4294967.295" step="0.01" style="width:76px"></label></div><div class="ribbon-stack"><span>Advance Slide</span><label class="ribbon-check"><input id="advance-click" type="checkbox"> On Mouse Click</label><label class="ribbon-check"><input id="advance-after" type="checkbox"> After: <input id="advance-time" type="text" aria-label="Advance slide after" value="00:00.00" inputmode="decimal" style="width:76px"></label></div>${command('transition-apply-all', 'Apply To All', 'normal', 'data-edit="transition-apply-all"')}<span class="group-label">Timing</span></div></div>
<div class="ribbon" id="ribbon-show" role="tabpanel" aria-labelledby="tab-show" hidden><div class="ribbon-group">${command('present-first', 'Play from<br>Start', 'first')}${command('present-current', 'Play from<br>Current Slide', 'play')}${command('presenter-view-start', 'Presenter<br>View', 'normal')}${command('custom-shows', 'Custom Show ▾', 'play', 'data-edit="custom-shows" aria-haspopup="menu"')}<span class="group-label">Start Slide Show</span></div><div class="ribbon-group">${command('show-properties', 'Set Up<br>Slide Show', 'normal', 'data-edit="show-properties"')}<span class="group-label">Set Up</span></div></div>
<div class="ribbon" id="ribbon-view" role="tabpanel" aria-labelledby="tab-view" hidden><div class="ribbon-group">${command('view-normal', 'Normal', 'normal', 'aria-pressed="true"')}<button id="view-sorter" class="ribbon-command" data-edit="sorter" aria-pressed="false">${officeIcon('normal')}<span>Slide Sorter</span></button><span class="group-label">Presentation Views</span></div><div class="ribbon-group"><label class="ribbon-check"><input id="show-thumbnails" type="checkbox" checked> Thumbnails</label><label class="ribbon-check"><input id="show-guides" type="checkbox"> Guides</label><button id="view-grid-guides" aria-haspopup="menu">Grid and Guides ▾</button><span class="group-label">Show</span></div><div class="ribbon-group">${command('view-zoom', 'Zoom', 'zoom')}${command('view-fit', 'Fit to<br>Window', 'fit')}<span class="group-label">Zoom</span></div></div>
<button id="collapse-ribbon" aria-label="Collapse ribbon" aria-expanded="true" title="Collapse ribbon">⌃</button>
</header>`;
export const officeFooter = `<footer><span id="count" aria-live="polite">No slides</span><span class="footer-spacer"></span><button data-edit="notes">Notes</button><button id="status-agents" title="Agents">${officeIcon('agent')} Agents</button><button id="status-normal" aria-label="Normal view" title="Normal view" aria-pressed="true">${officeIcon('normal')}</button><button id="status-present" aria-label="Start slide show" title="Slide Show">${officeIcon('play')}</button><div class="zoom-controls"><button id="zoom-out" aria-label="Zoom out">−</button><input id="zoom-slider" type="range" min="0" max="2000" value="1000" aria-label="Zoom percentage"><button id="zoom-in" aria-label="Zoom in">+</button><button id="zoom-level" title="Zoom">100%</button><button id="zoom-fit" aria-label="Fit slide to window" title="Fit slide to current window">${officeIcon('fit')}</button></div><select id="zoom" aria-label="Zoom" hidden><option value="fit">Fit</option><option value="0.5">50%</option><option value="0.75">75%</option><option value="1">100%</option><option value="1.25">125%</option><option value="1.5">150%</option><option value="2">200%</option></select></footer>
<dialog id="grid-dialog" aria-labelledby="grid-title"><form method="dialog"><h2 id="grid-title">Grid and Guides</h2><fieldset><legend>Snap to</legend><label><input id="grid-snap" type="checkbox">Snap objects to grid</label></fieldset><fieldset><legend>Grid Settings</legend><label>Spacing: <select id="grid-spacing">${[8, 6, 5, 4, 3, 2].map((n) => `<option value="${360000 / n}">${n} grids per cm</option>`).join('')}${[1, 2, 3, 4, 5].map((n) => `<option value="${360000 * n}">${n}cm</option>`).join('')}<option value="custom">Custom</option></select></label><label id="grid-custom-row" hidden>Grid spacing every <input id="grid-custom" type="number" min="0.001" max="100" step="any" aria-label="Grid spacing in centimeters"> centimeters</label><label><input id="grid-visible" type="checkbox">Display grid on screen</label></fieldset><fieldset><legend>Guide Settings</legend><label><input id="grid-guides" type="checkbox">Display drawing guides on screen</label><label><input id="grid-smart" type="checkbox">Display smart guides when shapes are aligned</label></fieldset><div class="dialog-actions"><button id="grid-default" type="button">Set as Default</button><button value="cancel" formnovalidate>Cancel</button><button value="apply">OK</button></div></form></dialog>
<dialog id="zoom-dialog" aria-labelledby="zoom-title"><form method="dialog"><h2 id="zoom-title">Zoom</h2><fieldset class="zoom-presets" aria-label="Zoom to">${['fit', '400', '200', '100', '66', '50', '33'].map((value) => `<label><input type="radio" name="zoom-preset" value="${value}">${value === 'fit' ? 'Fit' : value + '%'}</label>`).join('')}</fieldset><label class="zoom-custom">Percent: <input id="zoom-percent" type="number" min="10" max="400" value="100" required></label><div class="dialog-actions"><button value="cancel" formnovalidate>Cancel</button><button value="apply">OK</button></div></form></dialog>
<div id="office-menu" class="office-menu" role="menu" hidden></div>
<div id="show-screen" hidden></div>`;
