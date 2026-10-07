// The shell and the agent panes sit next to the editor, so they share its
// design tokens. These values are copied from packages/editor/src/ui/tokens.css
// (the shell is plain HTML and cannot import the editor's stylesheet);
// test/shell-tokens.test.mjs fails when the two drift apart.
export const tokenStyles = `:root{
--ok-accent:#b4235a;--ok-accent-2:#d6336c;--ok-bg:#f4f5f7;--ok-panel:#ffffff;--ok-panel-2:#f9fafb;--ok-ribbon:#f7f8fa;--ok-ribbon-active:#ffffff;
--ok-border:#dfe2e7;--ok-border-strong:#c9cdd6;--ok-text:#15171c;--ok-text-2:#545b69;--ok-text-3:#8a909c;--ok-hover:#eef0f3;--ok-selected:#fdebf1;--ok-selected-border:#b4235a;
--ok-canvas-bg:#d3d6dc;--ok-danger:#c92a2a;--ok-radius:3px;--ok-radius-lg:6px;
--ok-shadow:0 1px 2px rgba(16, 18, 23, 0.08), 0 0 0 1px rgba(16, 18, 23, 0.04);--ok-shadow-lg:0 12px 32px -8px rgba(16, 18, 23, 0.22), 0 2px 6px rgba(16, 18, 23, 0.08);
--ok-font:system-ui, -apple-system, 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;--ok-mono:ui-monospace, 'SF Mono', 'Cascadia Code', 'Consolas', monospace;
--shell-alert-bg:#fff0cd;--shell-alert-text:#583b00;
color-scheme:light;
}
@media (prefers-color-scheme: dark){:root{
color-scheme:dark;
--ok-accent:#e85a90;--ok-accent-2:#ff9cbf;--ok-bg:#1f1f1f;--ok-panel:#2b2b2b;--ok-panel-2:#252525;--ok-ribbon:#2a2a2a;--ok-ribbon-active:#333333;
--ok-border:#3d3d3d;--ok-border-strong:#555555;--ok-text:#f0f0f0;--ok-text-2:#c4c4c4;--ok-text-3:#8f8f8f;--ok-hover:#3a3a3a;--ok-selected:#3a2230;--ok-selected-border:#ff7aa8;
--ok-canvas-bg:#1b1b1b;--ok-danger:#f1707a;
--ok-shadow:0 1px 2px rgba(0, 0, 0, 0.45), 0 0 0 1px rgba(255, 255, 255, 0.04);--ok-shadow-lg:0 12px 32px -8px rgba(0, 0, 0, 0.6), 0 2px 6px rgba(0, 0, 0, 0.35);
--shell-alert-bg:#4a3a12;--shell-alert-text:#ffe2a3;
}}
*{box-sizing:border-box}
body{margin:0;background:var(--ok-bg);color:var(--ok-text);font:13px var(--ok-font);-webkit-font-smoothing:antialiased}
button,a,select,textarea{font:inherit;color:inherit}
button,select,.download{border:1px solid transparent;border-radius:var(--ok-radius);background:transparent;padding:3px 8px;text-decoration:none;cursor:pointer;white-space:nowrap}
button:hover:not(:disabled),.download:hover{background:var(--ok-hover);border-color:var(--ok-border)}
button:active:not(:disabled){background:var(--ok-selected)}
button[aria-pressed="true"]{background:var(--ok-selected);border-color:var(--ok-selected-border)}
button:disabled{color:var(--ok-text-3);cursor:default}
select{border-color:var(--ok-border-strong);background:var(--ok-panel);padding:2px 6px}
select:hover:not(:disabled){border-color:var(--ok-border-strong)}
:focus-visible{outline:2px solid var(--ok-selected-border);outline-offset:1px}
.primary{background:var(--ok-accent);border-color:var(--ok-accent);color:#fff}
.primary:hover:not(:disabled){background:var(--ok-accent-2);border-color:var(--ok-accent-2)}
.primary:disabled{background:var(--ok-hover);border-color:var(--ok-border);color:var(--ok-text-3)}
[hidden]{display:none!important}
@media(prefers-reduced-motion:reduce){*{scroll-behavior:auto!important;transition:none!important}}
`;

export const previewStyles = `${tokenStyles}
body{height:100dvh;overflow:hidden;display:grid;grid-template-rows:minmax(0,1fr) 26px}
.workspace{--filmstrip-width:190px;--chat-width:38vw;display:grid;grid-template-columns:var(--filmstrip-width) minmax(0,1fr) clamp(280px,var(--chat-width),calc(100vw - var(--filmstrip-width) - 240px));min-height:0}
.filmstrip{background:var(--ok-panel-2);border-right:1px solid var(--ok-border);overflow:auto;overscroll-behavior:contain;padding:8px}
.filmstrip h2{margin:4px 0 8px 26px;font-size:11px;font-weight:600;color:var(--ok-text-2)}
#thumbnails{display:flex;flex-direction:column;gap:6px;margin:0;padding:0;list-style:none}
.thumbnail{display:flex;align-items:flex-start;gap:6px;width:100%;padding:6px 6px 8px 4px;border:0;background:transparent;border-radius:0;text-align:left}
.thumbnail:hover:not(:disabled){background:transparent;border-color:transparent}
.slide-number{width:16px;flex:none;text-align:right;font-size:11px;color:var(--ok-text-2);padding-top:2px}
.thumbnail img{display:block;min-width:0;width:calc(100% - 22px);background:white;border:1px solid var(--ok-border);border-radius:4px;aspect-ratio:var(--slide-ratio,16/9);object-fit:contain}
.thumbnail:hover img{outline:3px solid var(--ok-border-strong);outline-offset:2px}
.thumbnail[aria-current="true"] img{outline:3px solid var(--ok-accent);outline-offset:2px}.thumbnail[aria-current="true"] .slide-number{color:var(--ok-accent);font-weight:600}
.thumbnail[data-skipped="true"] .slide-number{text-decoration:line-through}
.thumbnail[data-skipped="true"] img{opacity:.6}
main{min-width:0;min-height:0;display:flex;flex-direction:column}
#error{flex:none;max-height:30%;overflow:auto;background:var(--shell-alert-bg);color:var(--shell-alert-text);margin:0;padding:8px 12px;font:12px var(--ok-mono);white-space:pre-wrap;border-bottom:1px solid var(--ok-border)}
#stage{background:var(--ok-canvas-bg);flex:1;min-height:0;overflow:auto;display:flex;padding:32px;overscroll-behavior:contain}
#slide{position:relative;flex:none;margin:auto;background:white;box-shadow:var(--ok-shadow-lg);overflow:hidden;user-select:text;-webkit-user-select:text}
#empty{margin:auto;color:var(--ok-text-2)}
footer{display:flex;align-items:center;gap:6px;padding:0 12px;background:var(--ok-ribbon);border-top:1px solid var(--ok-border);color:var(--ok-text-2);font-size:11px;min-width:0;white-space:nowrap}
footer button,footer .download,footer select{font-size:11px;padding:1px 6px;line-height:16px}
footer .hint{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;color:var(--ok-text-3)}
#status{min-width:0;overflow:hidden;text-overflow:ellipsis}
.status-views{display:flex;gap:2px}
footer svg{display:block;fill:none;stroke:currentColor;stroke-width:1.1}
#toggle-editor{padding:2px 6px}
#toggle-chat svg{width:13px;height:13px}#toggle-chat[aria-pressed="true"] svg{fill:var(--ok-accent);stroke:var(--ok-accent)}#toggle-chat{display:inline-flex;align-items:center;gap:4px}
#notice{position:fixed;z-index:40;left:50%;bottom:36px;transform:translateX(-50%);max-width:min(480px,calc(100vw - 32px));padding:8px 12px;border:1px solid var(--ok-border);border-radius:var(--ok-radius-lg);background:var(--ok-panel);color:var(--ok-text);box-shadow:var(--ok-shadow-lg);font-size:12px}
#presentation-controls{display:none}
body.presenting{grid-template-rows:minmax(0,1fr);background:#111}
.presenting footer,.presenting .filmstrip,.presenting #chat{display:none}
.presenting .workspace{grid-template-columns:minmax(0,1fr)}.presenting #stage{padding:0;background:#111}.presenting #slide{box-shadow:none}
.presenting #presentation-controls{display:flex;position:fixed;bottom:16px;left:50%;transform:translateX(-50%);align-items:center;gap:12px;background:#202735e8;color:white;padding:6px;border-radius:8px;opacity:0;transition:opacity .15s}
.presenting #presentation-controls:hover,.presenting #presentation-controls:focus-within{opacity:1}
.presenting.browse-scrollbar #stage{padding-right:18px}
#presentation-scrollbar{position:fixed;z-index:4;top:0;bottom:0;right:0;width:16px;touch-action:none;border:1px solid #5d6575;background:#202735e8;cursor:pointer}
#presentation-scrollbar span{position:absolute;left:3px;right:3px;min-height:8px;border-radius:5px;background:#c4c4c4;pointer-events:none}
.kiosk-presenting #present-prev,.kiosk-presenting #present-next{display:none}
#presentation-controls button{background:transparent;color:white;border-color:#5d6575;padding:5px 10px}
#presentation-controls button:hover:not(:disabled){background:#ffffff1f}
#present-note{max-width:38ch;color:#ffd9a8;font-size:11px;line-height:1.35}
#chat{position:relative;min-width:0;min-height:0;display:flex;flex-direction:column;background:var(--ok-panel);color:var(--ok-text);border-left:1px solid var(--ok-border)}
#chat-resizer{position:absolute;left:-4px;top:0;bottom:0;width:7px;z-index:3;cursor:col-resize;touch-action:none}#chat-resizer::after{content:"";position:absolute;left:3px;top:0;bottom:0;width:1px}#chat-resizer:hover::after,#chat-resizer:focus-visible::after,.resizing-chat #chat-resizer::after{background:var(--ok-selected-border)}#chat-resizer:focus-visible{outline:none}
body.resizing-chat,body.resizing-chat *{cursor:col-resize!important;user-select:none!important}
.pane-head{display:flex;align-items:center;gap:8px;padding:10px 12px;border-bottom:1px solid var(--ok-border);background:var(--ok-panel);flex:none}
.pane-head strong{font-size:13px;font-weight:600}
.pane-close{margin-left:auto;padding:0 4px;border:none;background:none;color:var(--ok-text-2);font-size:20px;line-height:20px}
.pane-close:hover:not(:disabled){background:var(--ok-hover);border-color:transparent;color:var(--ok-text)}
body.chat-hidden .workspace{grid-template-columns:var(--filmstrip-width) minmax(0,1fr)}body.chat-hidden #chat{display:none}
body.presenting .workspace{grid-template-columns:minmax(0,1fr)}
@media(max-width:1000px){.workspace{--filmstrip-width:120px;--chat-width:380px}#stage{padding:16px}}
@media(max-width:700px){#chat-resizer{display:none}.workspace{grid-template-columns:minmax(0,1fr);grid-template-rows:minmax(160px,1fr) minmax(200px,1fr)}.filmstrip{display:none}#chat{border-left:0;border-top:1px solid var(--ok-border)}body.chat-hidden .workspace,.presenting .workspace{grid-template-columns:minmax(0,1fr);grid-template-rows:minmax(0,1fr)}footer .hint,#status,.download{display:none}#stage{padding:16px}}
`;

/** The agent page renders inside one agent pane of the docked Agents task pane. */
export const agentStyles = `${tokenStyles}
body{display:flex;flex-direction:column;background:var(--ok-panel);height:100dvh;min-width:0;overflow:auto;overflow-x:hidden}
.chat-heading{display:flex;align-items:center;flex-wrap:wrap;gap:6px;padding:6px 12px 2px;flex-shrink:0}
.chat-heading select{flex:1;min-width:100px;font-size:12px}
#chat-context{flex-shrink:0;margin:0;padding:4px 12px 6px;color:var(--ok-text-2);font-size:11px;border-bottom:1px solid var(--ok-border)}
#claude-panel,#codex-panel{display:flex;flex-direction:column;flex:1;min-height:0}
#claude-panel{min-height:100px}#codex-panel{min-height:240px}
.terminal-toolbar{display:flex;align-items:center;gap:6px;padding:6px 12px;min-height:34px;font-size:11px;color:var(--ok-text-2)}
.terminal-toolbar span{flex:1;overflow-wrap:anywhere}
.terminal-toolbar button{font-size:11px;padding:2px 8px;border-color:var(--ok-border-strong)}
.terminal-frame{display:flex;flex:1;min-height:0;min-width:0;padding:2px 4px 4px 12px;overflow:hidden}
#terminal{flex:1;width:100%;min-width:0;min-height:0;overflow:hidden;padding:0}.xterm{height:100%}
/* xterm's stylesheet paints the viewport black; the rows below the last full line show it. */
#terminal .xterm-viewport{background-color:var(--ok-panel)!important}
#messages{flex:1;overflow:auto;padding:12px;display:flex;flex-direction:column;gap:12px;overscroll-behavior:contain}
.chat-message{white-space:pre-wrap;overflow-wrap:anywhere;line-height:1.55}.chat-message.user{background:var(--ok-panel-2);border:1px solid var(--ok-border);border-radius:var(--ok-radius-lg);padding:8px 10px}
.chat-message small{display:block;color:var(--ok-text-2);font-size:11px;margin-bottom:4px}.chat-intro{color:var(--ok-text-2);line-height:1.6;font-size:12px}
#chat-form{padding:8px 12px;border-top:1px solid var(--ok-border)}
#chat-input{resize:vertical;min-height:70px;max-height:200px;width:100%;padding:6px 8px;border:1px solid var(--ok-border-strong);border-radius:var(--ok-radius);background:var(--ok-panel);color:var(--ok-text)}
#chat-input:focus{outline:2px solid var(--ok-selected-border);outline-offset:-1px}
.chat-actions{display:flex;gap:6px;margin-top:6px;align-items:center}.chat-actions button{border-color:var(--ok-border-strong)}.chat-actions .primary{border-color:var(--ok-accent)}
.chat-shortcut{flex:1;color:var(--ok-text-3);font-size:10px}
#chat-status{padding:0 12px 8px;color:var(--ok-text-2);font-size:11px;white-space:pre-wrap;max-height:90px;overflow:auto}
.markdown{white-space:normal;line-height:1.6}.markdown>:first-child{margin-top:0}.markdown>:last-child{margin-bottom:0}.markdown p{margin:0 0 .8em}
.markdown pre{overflow:auto;padding:8px;background:var(--ok-panel-2);border:1px solid var(--ok-border);border-radius:var(--ok-radius);white-space:pre;font-family:var(--ok-mono)}
.markdown code{font:.9em var(--ok-mono);background:var(--ok-panel-2);border-radius:3px;padding:1px 4px}.markdown pre code{padding:0}
.markdown ul,.markdown ol{padding-left:22px}.markdown blockquote{margin-left:0;padding-left:12px;border-left:3px solid var(--ok-border);color:var(--ok-text-2)}
.markdown table{display:block;overflow:auto;border-collapse:collapse}.markdown th,.markdown td{padding:6px 8px;border:1px solid var(--ok-border)}
.markdown h1,.markdown h2,.markdown h3{font-size:1.1em}.markdown a{color:var(--ok-accent);text-decoration:underline}
`;
