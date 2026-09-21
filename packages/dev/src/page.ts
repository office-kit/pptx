import { previewStyles } from './styles.ts';
export const page = `<!doctype html>
<html lang="en">
<meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>Office Kit — PowerPoint preview</title>
<link rel="stylesheet" href="/terminal.css">
<style>
${previewStyles}
#agent-workspace{position:relative;flex:1;min-height:0;min-width:0;display:flex;overflow:auto}.agent-pane{position:absolute;padding:3px 5px;display:flex;flex-direction:column;flex:1;min-width:0;min-height:0;overflow:hidden}.agent-pane iframe{border:0;width:100%;flex:1;min-height:0;border-radius:0 0 9px 9px}.agent-tools{display:flex;gap:4px;align-items:center;flex-wrap:wrap;padding:6px;background:#20253a;border:1px solid #34334e;border-radius:9px 9px 0 0}.agent-tools span{flex:1;font-size:10px;letter-spacing:.04em;color:#b3abc9}.agent-tools svg{width:14px;height:14px}.agent-pane:focus-within .agent-tools{border-color:#8b6ed577;background:#29243e}.agent-tools button{font-size:11px;padding:5px;background:transparent;border:0;display:grid;place-items:center}.agent-divider{position:absolute;z-index:2;background:#101422;cursor:col-resize;touch-action:none}.agent-divider.vertical{cursor:row-resize}.agent-divider:hover,.agent-divider:focus-visible{background:#9b87ff}.splitting-agents iframe,.resizing-chat iframe{pointer-events:none}
#editor-frame{display:none;border:0;width:100%;height:100%;flex:1;min-height:0}
body.editing:not(.presenting){grid-template-rows:60px minmax(0,1fr)}
.editing:not(.presenting) .workspace{--filmstrip-width:0px;grid-template-columns:minmax(0,1fr) clamp(280px,var(--chat-width),calc(100vw - 500px))}
.editing.chat-hidden:not(.presenting) .workspace{grid-template-columns:minmax(0,1fr)}
.editing:not(.presenting) .filmstrip,.editing:not(.presenting) footer,.editing:not(.presenting) #stage,.editing:not(.presenting) #present,.editing:not(.presenting) .download{display:none}
.editing:not(.presenting) #editor-frame{display:block}
@media(max-width:900px){.editing:not(.presenting) .workspace{grid-template-columns:minmax(0,1fr)}.editing:not(.presenting) #chat{display:none}}
</style>
<body class="editing"><header><span class="brand"><svg viewBox="0 0 36 36" aria-hidden="true"><defs><linearGradient id="brand-gradient" x2="1" y2="1"><stop stop-color="#c39aff"/><stop offset="1" stop-color="#7165f4"/></linearGradient></defs><rect x="2" y="2" width="32" height="32" rx="10" fill="url(#brand-gradient)"/><path d="M11 10h9l6 6v10H11z" fill="none" stroke="white" stroke-width="1.8" stroke-linejoin="round"/><path d="M20 10v7h6M15 21h7M15 25h4" fill="none" stroke="white" stroke-width="1.8" stroke-linecap="round"/><path d="m28 3 1.2 3.8L33 8l-3.8 1.2L28 13l-1.2-3.8L23 8l3.8-1.2z" fill="#e7dbff"/></svg>Office <em>Kit</em></span><span class="badge">Studio</span><span id="status" role="status">Building…</span><button id="toggle-editor" aria-pressed="true">Preview</button><button id="toggle-chat" aria-expanded="true" aria-controls="chat">✦ Agents</button><button id="present" disabled>Present</button><a class="download" href="/deck.pptx">Download PPTX</a></header>
<div class="workspace">
<nav class="filmstrip" aria-label="Slides"><h2>Slides</h2><ol id="thumbnails"></ol></nav>
<main aria-label="Slide viewer"><iframe id="editor-frame" title="Presentation editor"></iframe><pre id="error" role="alert" hidden></pre><div id="stage" tabindex="-1"><div id="empty">Waiting for slides…</div><div id="slide" hidden></div></div></main>
<aside id="chat" aria-label="Slide chat">
<div id="chat-resizer" role="separator" tabindex="0" aria-label="Chat width" aria-orientation="vertical" aria-controls="chat" title="Drag to resize · Double-click to reset"></div>
<div class="workspace-heading"><span>✦ AI WORKSPACE</span><b>LOCAL</b></div><div id="agent-workspace"></div><div id="chat-context" hidden></div></aside>
</div>
<footer><span id="count" aria-live="polite">No slides</span><span class="hint">Changes appear automatically · Text can be selected and copied</span><button id="prev" aria-label="Previous slide" disabled>‹</button><button id="next" aria-label="Next slide" disabled>›</button><label for="zoom">Zoom</label><select id="zoom"><option value="fit">Fit</option><option value="0.5">50%</option><option value="0.75">75%</option><option value="1">100%</option><option value="1.25">125%</option><option value="1.5">150%</option><option value="2">200%</option></select></footer>
<div id="presentation-controls"><button id="present-prev" aria-label="Previous slide">‹</button><span id="present-count"></span><button id="present-next" aria-label="Next slide">›</button><button id="exit-present">Exit · Esc</button></div>
<script>
let state={slides:[],error:null,aspectRatio:16/9},index=0,urls=[],presenting=false;
let displayedSvg;
const byId=id=>document.getElementById(id);
const stage=byId('stage'),slide=byId('slide'),thumbnails=byId('thumbnails');
const workspace=document.querySelector('.workspace'),chatResizer=byId('chat-resizer');
const chatWidthKey='office-kit-chat-width',minChatWidth=280,minStageWidth=240;
function chatWidthLimits(){return {min:minChatWidth,max:Math.max(minChatWidth,workspace.clientWidth-document.querySelector('.filmstrip').getBoundingClientRect().width-minStageWidth)};}
function updateChatWidthAria(){
  const {min,max}=chatWidthLimits();
  chatResizer.setAttribute('aria-valuemin',String(min));chatResizer.setAttribute('aria-valuemax',String(Math.round(max)));
  const width=Math.round(byId('chat').getBoundingClientRect().width);
  chatResizer.setAttribute('aria-valuenow',String(width));chatResizer.setAttribute('aria-valuetext',width+' pixels');
}
function setChatWidth(width){const {min,max}=chatWidthLimits();workspace.style.setProperty('--chat-width',Math.min(max,Math.max(min,width))+'px');}
function saveChatWidth(){try{const width=workspace.style.getPropertyValue('--chat-width');if(width)localStorage.setItem(chatWidthKey,String(parseFloat(width)));else localStorage.removeItem(chatWidthKey);}catch(error){console.warn('Could not save chat width',error);}}
try{const saved=Number(localStorage.getItem(chatWidthKey));if(Number.isFinite(saved)&&saved>=minChatWidth)workspace.style.setProperty('--chat-width',saved+'px');}catch(error){console.warn('Could not restore chat width',error);}
let chatDrag=null;
chatResizer.onpointerdown=event=>{
  if(event.button!==0||!event.isPrimary)return;
  event.preventDefault();chatResizer.focus();chatDrag={id:event.pointerId,x:event.clientX,width:byId('chat').getBoundingClientRect().width};
  chatResizer.setPointerCapture(event.pointerId);document.body.classList.add('resizing-chat');
};
chatResizer.onpointermove=event=>{if(chatDrag?.id===event.pointerId)setChatWidth(chatDrag.width+chatDrag.x-event.clientX);};
function finishChatResize(){if(!chatDrag)return;chatDrag=null;document.body.classList.remove('resizing-chat');saveChatWidth();}
chatResizer.onpointerup=finishChatResize;chatResizer.onpointercancel=finishChatResize;chatResizer.onlostpointercapture=finishChatResize;
chatResizer.ondblclick=()=>{workspace.style.removeProperty('--chat-width');saveChatWidth();};
chatResizer.onkeydown=event=>{
  const {min,max}=chatWidthLimits(),width=byId('chat').getBoundingClientRect().width,step=event.shiftKey?50:10;
  let next;
  if(event.key==='ArrowLeft')next=width+step;else if(event.key==='ArrowRight')next=width-step;else if(event.key==='Home')next=min;else if(event.key==='End')next=max;else return;
  event.preventDefault();event.stopPropagation();setChatWidth(next);saveChatWidth();
};
new ResizeObserver(updateChatWidthAria).observe(byId('chat'));

// The slide lives in a shadow root rather than a sandboxed iframe so its text
// (XHTML inside <foreignObject>) can be selected and copied while the deck's
// SVG stays out of the viewer's own DOM and CSS. Keyboard events still reach the
// document, so arrow-key navigation keeps working after clicking into a slide.
const canvas=slide.attachShadow({mode:'open'});
function resize(){
  if(!state.slides.length)return;
  const style=getComputedStyle(stage);
  const width=Math.max(1,stage.clientWidth-parseFloat(style.paddingLeft)-parseFloat(style.paddingRight));
  const height=Math.max(1,stage.clientHeight-parseFloat(style.paddingTop)-parseFloat(style.paddingBottom));
  const ratio=state.aspectRatio;
  const zoom=byId('zoom').value;
  const slideWidth=presenting||zoom==='fit'?Math.min(width,height*ratio):1280*Number(zoom);
  slide.style.width=slideWidth+'px';slide.style.height=slideWidth/ratio+'px';
}
function selectSlide(next,focusThumbnail=false,reveal=true){
  index=Math.max(0,Math.min(next,state.slides.length-1));
  const count=state.slides.length?'Slide '+(index+1)+' of '+state.slides.length:'No slides';
  byId('chat-context').textContent=state.slides.length?count:'Whole project';
  byId('chat-context').dataset.focus=JSON.stringify({slide:state.slides.length?index:null,revision:state.revision??0});
  window.dispatchEvent(new Event('agent-focus'));
  if(document.body.classList.contains('editing')&&editorFocus)applyEditorFocus();
  byId('count').textContent=count;byId('present-count').textContent=count;
  for(const id of ['prev','present-prev'])byId(id).disabled=index===0;
  for(const id of ['next','present-next'])byId(id).disabled=index>=state.slides.length-1;
  byId('present').disabled=!state.slides.length;
  byId('zoom').disabled=!state.slides.length;
  slide.hidden=!state.slides.length;byId('empty').hidden=!!state.slides.length;
  const svg=state.slides[index];
  if(svg!==displayedSvg){
    displayedSvg=svg;
    canvas.innerHTML=svg?'<style>svg{display:block;width:100%;height:100%}</style>'+svg:'';
  }
  slide.setAttribute('aria-label','Slide '+(index+1));
  for(const [position,item] of Array.from(thumbnails.children).entries()){
    const button=item.firstElementChild,selected=position===index;
    button.setAttribute('aria-current',String(selected));button.tabIndex=selected?0:-1;
    if(selected){if(reveal)button.scrollIntoView({block:'nearest'});if(focusThumbnail)button.focus({preventScroll:true});}
  }
  resize();
}
function update(updated){
  const focusedThumbnail=thumbnails.contains(document.activeElement);
  const previous=state;
  state=updated;
  byId('status').textContent=state.error?'Build failed · showing last successful output':state.building?'Updating…':state.slides.length+' slides · Live';
  byId('error').textContent=state.error||'';byId('error').hidden=!state.error;
  document.documentElement.style.setProperty('--slide-ratio',String(state.aspectRatio));
  for(let i=state.slides.length;i<urls.length;i++){
    URL.revokeObjectURL(urls[i]);thumbnails.lastElementChild.remove();
  }
  urls.length=state.slides.length;
  state.slides.forEach((svg,i)=>{
    if(svg===previous.slides[i])return;
    const oldUrl=urls[i];
    urls[i]=URL.createObjectURL(new Blob([svg],{type:'image/svg+xml'}));
    let item=thumbnails.children[i];
    if(!item){
      item=document.createElement('li');
      const button=document.createElement('button'),number=document.createElement('span'),image=document.createElement('img');
      button.className='thumbnail';button.setAttribute('aria-label','Slide '+(i+1));button.onclick=()=>selectSlide(i,true);
      number.className='slide-number';number.textContent=String(i+1);
      image.alt='';image.draggable=false;image.loading='lazy';
      button.append(number,image);item.append(button);thumbnails.append(item);
    }
    item.querySelector('img').src=urls[i];
    if(oldUrl)URL.revokeObjectURL(oldUrl);
  });
  selectSlide(index,focusedThumbnail,false);
  if(!state.slides.length&&presenting)void exitPresentation();
}
function setPresenting(value){
  presenting=value;document.body.classList.toggle('presenting',value);resize();
  if(value)stage.focus();else{
    byId('present').focus();
    thumbnails.children[index]?.firstElementChild.scrollIntoView({block:'nearest'});
  }
}
async function exitPresentation(){
  setPresenting(false);
  if(document.fullscreenElement)await document.exitFullscreen();
}
byId('present').onclick=async()=>{
  setPresenting(true);
  try{await document.documentElement.requestFullscreen();}
  catch{byId('exit-present').textContent='Exit view · Esc';}
};
byId('exit-present').onclick=exitPresentation;
document.addEventListener('fullscreenchange',()=>{if(!document.fullscreenElement&&presenting)setPresenting(false);});
for(const id of ['prev','present-prev'])byId(id).onclick=()=>selectSlide(index-1);
for(const id of ['next','present-next'])byId(id).onclick=()=>selectSlide(index+1);
byId('zoom').onchange=resize;
stage.onclick=()=>{if(presenting&&!getSelection().toString())selectSlide(index+1);};
document.addEventListener('keydown',event=>{
  if(event.key==='Escape'&&presenting){event.preventDefault();void exitPresentation();return;}
  if(event.altKey||event.ctrlKey||event.metaKey||event.target.closest('select,input,textarea,[contenteditable]'))return;
  let next=index;
  if(['ArrowLeft','ArrowUp','PageUp'].includes(event.key))next--;
  else if(['ArrowRight','ArrowDown','PageDown'].includes(event.key))next++;
  else if(event.key==='Home')next=0;
  else if(event.key==='End')next=state.slides.length-1;
  else if(event.key===' '&&presenting&&!event.target.closest('button,a'))next+=event.shiftKey?-1:1;
  else return;
  event.preventDefault();selectSlide(next,thumbnails.contains(document.activeElement));
});
new ResizeObserver(resize).observe(stage);
let refreshId=0;
async function refresh(){
  const id=++refreshId;
  try{const response=await fetch('/state'+(state.revision===undefined?'':'?since='+state.revision));if(!response.ok)throw new Error('Preview unavailable');const updated=await response.json();if(id===refreshId){if(updated.changes){updated.slides=state.slides.slice(0,updated.count);updated.slides.length=updated.count;for(const [position,svg] of Object.entries(updated.changes))updated.slides[Number(position)]=svg;}update(updated);}}
  catch{if(id===refreshId)byId('status').textContent='Reconnecting…';}
}
byId('toggle-chat').onclick=()=>{const hidden=document.body.classList.toggle('chat-hidden');byId('toggle-chat').setAttribute('aria-expanded',String(!hidden));resize();};
const editorFrame=byId('editor-frame');
let editorFocus;
function setEditorMode(editing){
 document.body.classList.toggle('editing',editing);
 if(editing&&!editorFrame.getAttribute('src'))editorFrame.src='/editor';
 byId('toggle-editor').setAttribute('aria-pressed',String(editing));
 byId('toggle-editor').textContent=editorFocus?.locale==='ja'?(editing?'プレビュー':'編集'):(editing?'Preview':'Edit');
 if(editing&&editorFocus)applyEditorFocus();else selectSlide(index);
 resize();
}
byId('toggle-editor').onclick=()=>{
 const editing=!document.body.classList.contains('editing');setEditorMode(editing);
 try{sessionStorage.setItem('office-kit-view',editing?'editor':'preview');}catch(error){console.warn('Could not save view preference',error);}
};
try{setEditorMode(sessionStorage.getItem('office-kit-view')!=='preview');}catch(error){console.warn('Could not restore view preference',error);setEditorMode(true);}
function applyEditorFocus(){
 if(!editorFocus)return;
 index=Math.max(0,editorFocus.slide);
 byId('chat-context').textContent=(editorFocus.locale==='ja'?'スライド ':'Slide ')+(index+1)+' / '+editorFocus.count+(editorFocus.dirty?' · '+(editorFocus.locale==='ja'?'未保存':'Unsaved'):'');
 byId('chat-context').dataset.focus=JSON.stringify({slide:editorFocus.count?index:null,revision:editorFocus.dirty?-1:editorFocus.revision});
 window.dispatchEvent(new Event('agent-focus'));
}
window.addEventListener('message',event=>{
 if(event.origin!==location.origin||event.source!==editorFrame.contentWindow||event.data?.type!=='editor-focus')return;
 editorFocus=event.data;
 if(document.body.classList.contains('editing'))applyEditorFocus();
 byId('toggle-editor').textContent=editorFocus.locale==='ja'?(document.body.classList.contains('editing')?'プレビュー':'編集'):(document.body.classList.contains('editing')?'Preview':'Edit');
});
const events=new EventSource('/events');events.onmessage=event=>{if(event.data==='chat'){window.dispatchEvent(new Event('agent-chat'));return;}if(event.data==='ready'){delete state.revision;}void refresh();};events.onerror=()=>{byId('status').textContent='Reconnecting…'};refresh();
</script><script type="module" src="/terminal.js"></script></html>`;
