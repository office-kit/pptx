export const previewStyles = `*{box-sizing:border-box}
body{margin:0;height:100dvh;overflow:hidden;background:#e9eaf4;color:#dce1f1;font:13px system-ui;display:grid;grid-template-rows:60px minmax(0,1fr) 42px}
button,a,select{font:inherit;color:inherit}button,select,.download{border:1px solid #30364d;border-radius:9px;background:#191e30;padding:7px 12px;text-decoration:none;cursor:pointer}
button:hover:not(:disabled),.download:hover{background:#292f48}button:disabled{opacity:.4;cursor:default}
:focus-visible{outline:2px solid #9b87ff;outline-offset:3px}
header{display:flex;align-items:center;gap:16px;padding:0 20px;background:linear-gradient(110deg,#171b2f,#211c38 60%,#171b2f);border-bottom:1px solid #38304d;min-width:0}
.brand{font-weight:700;font-size:17px;white-space:nowrap}.badge{font-size:11px;color:#b2a4fb;background:#342c54;border-radius:4px;padding:3px 6px}
#status{flex:1;color:#9ba5be;min-width:0}#present{background:#7154e8;color:#191e30;border-color:#7154e8}
.workspace{--filmstrip-width:190px;--chat-width:38vw;display:grid;grid-template-columns:var(--filmstrip-width) minmax(0,1fr) clamp(280px,var(--chat-width),calc(100vw - var(--filmstrip-width) - 240px));min-height:0}
.filmstrip{background:#111625;border-right:1px solid #30364d;overflow:auto;overscroll-behavior:contain;padding:16px 12px}
.filmstrip h2{margin:0 0 12px 26px;text-transform:uppercase;letter-spacing:.1em;font-size:10px;color:#8490ac;font-weight:600}
#thumbnails{display:flex;flex-direction:column;gap:12px;margin:0;padding:0;list-style:none}
.thumbnail{display:flex;align-items:flex-start;gap:8px;width:100%;padding:4px 3px;border:0;background:transparent;border-radius:5px;text-align:left}
.slide-number{width:18px;flex:none;text-align:right;font-size:11px;color:#8995b0;padding-top:5px}
.thumbnail img{display:block;min-width:0;width:calc(100% - 26px);background:white;box-shadow:0 1px 4px #080b1933;border:2px solid transparent;border-radius:3px;aspect-ratio:var(--slide-ratio,16/9);object-fit:contain}
.thumbnail[aria-current="true"]{background:#31274e}.thumbnail[aria-current="true"] img{border-color:#9b87ff}.thumbnail[aria-current="true"] .slide-number{color:#b6a5ff;font-weight:700}
main{min-width:0;min-height:0;display:flex;flex-direction:column}
#error{flex:none;max-height:30%;overflow:auto;background:#fff0ef;color:#922e25;margin:0;padding:16px 20px;white-space:pre-wrap;border-bottom:1px solid #e8bbb7}
#error[hidden]{display:none}
#stage{background:radial-gradient(ellipse at 20% 15%,#ddd8f6aa,transparent 65%),radial-gradient(#a0a4bf33 1px,transparent 1px),#edeef6;background-size:auto,18px 18px,auto;flex:1;min-height:0;overflow:auto;display:flex;padding:32px;overscroll-behavior:contain}
#slide{position:relative;flex:none;margin:auto;background:white;box-shadow:0 20px 70px #28234825,0 2px 8px #28234815;overflow:hidden}
#slide{user-select:text;-webkit-user-select:text}
#empty{margin:auto;color:#8995b0}
footer{display:flex;align-items:center;gap:14px;padding:0 16px;background:#191e30;border-top:1px solid #30364d;font-size:12px}
#count{min-width:90px}footer .hint{flex:1;color:#8995b0}footer button{padding:3px 10px}footer select{padding:3px 8px}
#presentation-controls{display:none}
body.presenting{grid-template-rows:minmax(0,1fr);background:#111}
.presenting header,.presenting footer,.presenting .filmstrip,.presenting #chat{display:none}
.presenting .workspace{grid-template-columns:minmax(0,1fr)}.presenting #stage{padding:0;background:#111}.presenting #slide{box-shadow:none}
.presenting #presentation-controls{display:flex;position:fixed;bottom:16px;left:50%;transform:translateX(-50%);align-items:center;gap:12px;background:#202735e8;color:white;padding:6px;border-radius:8px;opacity:0;transition:opacity .15s}
.presenting #presentation-controls:hover,.presenting #presentation-controls:focus-within{opacity:1}
#presentation-controls button{background:transparent;color:white;border-color:#5d6575}
@media(max-width:700px){.workspace{grid-template-columns:140px minmax(0,1fr)}.filmstrip{padding:12px 5px}header{padding:0 12px;gap:10px}.badge,footer .hint{display:none}#stage{padding:16px}footer{gap:8px}#status{font-size:11px}.download{padding:7px 8px}}
#chat{position:relative;min-width:0;min-height:0;display:flex;flex-direction:column;background:#191e30;border-left:1px solid #30364d}
#chat-resizer{position:absolute;left:-5px;top:0;bottom:0;width:10px;z-index:3;cursor:col-resize;touch-action:none}#chat-resizer::after{content:"";position:absolute;left:4px;top:0;bottom:0;width:2px}#chat-resizer:hover::after,#chat-resizer:focus-visible::after,.resizing-chat #chat-resizer::after{background:#9b87ff}body.resizing-chat,body.resizing-chat *{cursor:col-resize!important;user-select:none!important}
.chat-heading{display:flex;align-items:center;gap:8px;padding:16px;border-bottom:1px solid #2a3046}.chat-heading strong{flex:1}.chat-heading button{padding:5px 8px}
#messages{flex:1;overflow:auto;padding:16px;display:flex;flex-direction:column;gap:16px;overscroll-behavior:contain}
.chat-message{white-space:pre-wrap;overflow-wrap:anywhere;line-height:1.6}.chat-message.user{background:#302752;border-radius:8px;padding:10px 12px}.chat-message small{display:block;color:#a4afc6;font-size:11px;margin-bottom:4px}.chat-intro{color:#8995b0;line-height:1.7}
#chat-form{padding:14px;border-top:1px solid #2a3046}#chat-context{font-size:11px;color:#9b87ff;margin-bottom:8px}#chat-input{font:inherit;resize:vertical;min-height:90px;max-height:200px;width:100%;padding:10px;border:1px solid #30364d;border-radius:8px} .chat-actions{display:flex;gap:8px;margin-top:8px;align-items:center}.chat-actions select{min-width:0;flex:1}#chat-send{background:#7154e8;color:white}#chat-status{padding:0 14px 12px;color:#9ba5be;font-size:11px;white-space:pre-wrap;max-height:90px;overflow:auto}
body.chat-hidden .workspace{grid-template-columns:190px minmax(0,1fr)}body.chat-hidden #chat{display:none}
body.presenting .workspace{grid-template-columns:minmax(0,1fr)}
@media(max-width:1000px){.workspace{--filmstrip-width:120px;--chat-width:380px}.filmstrip{padding:12px 5px}#stage{padding:16px}}
@media(max-width:700px){#chat-resizer{display:none}.workspace{grid-template-columns:minmax(0,1fr);grid-template-rows:minmax(160px,1fr) minmax(200px,1fr)}.filmstrip{display:none}#chat{border-top:1px solid #30364d}.chat-heading{padding:8px 12px}#chat-form{padding:8px 12px}#chat-input{min-height:50px}body.chat-hidden .workspace,.presenting .workspace{grid-template-columns:minmax(0,1fr);grid-template-rows:minmax(0,1fr)}.brand{font-size:14px}header{gap:6px}.download{font-size:11px}}
@media(prefers-reduced-motion:reduce){*{scroll-behavior:auto!important;transition:none!important}}

[hidden]{display:none!important}
.chat-heading{padding:10px 14px}.chat-heading select{flex:1;min-width:0;font-weight:600;border:0;padding-left:0}
#chat-context{margin:0;padding:8px 14px;border-bottom:1px solid #2a3046;color:#9ba5be;font-size:11px}
#claude-panel,#codex-panel{display:flex;flex-direction:column;flex:1;min-height:0}
#claude-panel{background:#171b24;color:#e4e8f0}
.terminal-toolbar{display:flex;align-items:center;gap:8px;padding:8px 12px;font-size:11px;min-height:40px}.terminal-toolbar span{flex:1;overflow-wrap:anywhere}.terminal-toolbar button{font-size:11px;padding:5px 8px;background:#262d3b;border-color:#45516a;color:#e4e8f0}
.terminal-toolbar button:hover:not(:disabled){background:#35415a}
#terminal{flex:1;min-height:0;overflow:hidden;padding:0}.xterm{height:100%}
#chat-input{min-height:70px;resize:vertical}.chat-shortcut{flex:1;color:#9098a7;font-size:10px}
.markdown{white-space:normal;line-height:1.65}.markdown>:first-child{margin-top:0}.markdown>:last-child{margin-bottom:0}.markdown p{margin:0 0 .8em}.markdown pre{overflow:auto;padding:10px;background:#22283b;border:1px solid #2a3046;border-radius:6px;white-space:pre}.markdown code{font-size:.9em;background:#22283b;border-radius:3px;padding:1px 4px}.markdown pre code{padding:0}.markdown ul,.markdown ol{padding-left:22px}.markdown blockquote{margin-left:0;padding-left:12px;border-left:3px solid #30364d;color:#9ba5be}.markdown table{display:block;overflow:auto;border-collapse:collapse}.markdown th,.markdown td{padding:6px 8px;border:1px solid #30364d}.markdown h1,.markdown h2,.markdown h3{font-size:1.1em}.markdown a{color:#b6a5ff;text-decoration:underline}

.brand{display:flex;align-items:center;gap:10px;letter-spacing:-.5px}.brand svg{width:32px;height:32px;flex:none}.brand em{font-style:normal;font-weight:400;color:#aca4ce}.badge{text-transform:uppercase;letter-spacing:.14em;font-size:9px;border:1px solid #61508866;border-radius:5px}#status{font-size:11px}#present,#chat-send{color:#fff;background:linear-gradient(120deg,#8061ed,#634de0);border-color:#a38aef55;box-shadow:0 3px 15px #7455e82a}#toggle-chat[aria-expanded="true"]{background:#35294f;border-color:#6f528b;color:#d6c5ff}.download{background:transparent;color:#bcc5dc}header button,header .download{font-size:12px}.filmstrip{background:linear-gradient(180deg,#161b2c,#111625)}.thumbnail{transition:background .15s}.thumbnail[aria-current="true"]{box-shadow:inset 0 0 0 1px #9980ef33}.thumbnail[aria-current="true"] img{box-shadow:0 0 16px #9567ed22}footer{background:#141929;color:#b9c2d8}footer .hint{font-size:10px}#count{font-size:11px;font-weight:600}#chat{background:#101422}.workspace-heading{display:flex;align-items:center;justify-content:space-between;padding:12px 14px;color:#c7bee5;font-size:10px;font-weight:600;letter-spacing:.12em}.workspace-heading b{font-size:9px;color:#89cdbc;font-weight:500;letter-spacing:.08em}.workspace-heading b::before{content:"";display:inline-block;width:5px;height:5px;background:#86d5bb;border-radius:50%;margin-right:6px;box-shadow:0 0 8px #86d5bb66}.chat-heading{background:#191e30}.chat-heading select{background:transparent;color:#e4e5f3;font-size:12px}.chat-heading option{background:#191e30}#chat-context{background:#1a1e30;color:#a59abb;font-size:10px}.terminal-toolbar{gap:6px;padding:8px;font-size:10px}.terminal-toolbar button{white-space:nowrap}#terminal-start{background:#6c52c9;border-color:#9678e9}.terminal-frame{display:flex;flex:1;min-height:0;min-width:0;padding:4px 8px 8px;overflow:hidden}#terminal{width:100%;min-width:0}#chat-input{color:#e2e5f3;background:#131827;border-color:#3b3f58}#chat-input:focus{outline:none;border-color:#9b87ff;box-shadow:0 0 0 3px #9b87ff18}.chat-message.user{border:1px solid #67509655}.chat-intro{font-size:12px}.markdown pre{border-color:#3a4059}.markdown code{color:#d2c6fa}#error{background:#38242b;color:#ffb8b0}
@media(max-width:700px){.brand svg{width:26px;height:26px}.brand{gap:6px}.workspace-heading{padding:7px 12px}header .download{font-size:10px;padding:7px}.badge{display:none}}
`;
