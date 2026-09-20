export const page = `<!doctype html>
<html lang="en">
<meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>Office Kit — PowerPoint preview</title>
<style>
*{box-sizing:border-box}
body{margin:0;height:100dvh;overflow:hidden;background:#e9ecf1;color:#202735;font:13px system-ui;display:grid;grid-template-rows:60px minmax(0,1fr) 42px}
button,a,select{font:inherit;color:inherit}button,select,.download{border:1px solid #d4d9e2;border-radius:6px;background:#fff;padding:7px 12px;text-decoration:none;cursor:pointer}
button:hover:not(:disabled),.download:hover{background:#f0f3f9}button:disabled{opacity:.4;cursor:default}
:focus-visible{outline:2px solid #4967dd;outline-offset:3px}
header{display:flex;align-items:center;gap:16px;padding:0 20px;background:#fff;border-bottom:1px solid #d4d9e2;min-width:0}
.brand{font-weight:700;font-size:17px;white-space:nowrap}.badge{font-size:11px;color:#616b7c;background:#f1f3f7;border-radius:4px;padding:3px 6px}
#status{flex:1;color:#667085;min-width:0}#present{background:#293c73;color:#fff;border-color:#293c73}
.workspace{display:grid;grid-template-columns:190px minmax(0,1fr) 340px;min-height:0}
.filmstrip{background:#f7f8fa;border-right:1px solid #d4d9e2;overflow:auto;overscroll-behavior:contain;padding:16px 12px}
.filmstrip h2{margin:0 0 12px 26px;text-transform:uppercase;letter-spacing:.1em;font-size:10px;color:#7c8596;font-weight:600}
#thumbnails{display:flex;flex-direction:column;gap:12px;margin:0;padding:0;list-style:none}
.thumbnail{display:flex;align-items:flex-start;gap:8px;width:100%;padding:4px 3px;border:0;background:transparent;border-radius:5px;text-align:left}
.slide-number{width:18px;flex:none;text-align:right;font-size:11px;color:#737d8e;padding-top:5px}
.thumbnail img{display:block;min-width:0;width:calc(100% - 26px);background:white;box-shadow:0 1px 4px #19212d18;border:2px solid transparent;border-radius:3px;aspect-ratio:var(--slide-ratio,16/9);object-fit:contain}
.thumbnail[aria-current="true"]{background:#e7edff}.thumbnail[aria-current="true"] img{border-color:#4967dd}.thumbnail[aria-current="true"] .slide-number{color:#3654c1;font-weight:700}
main{min-width:0;min-height:0;display:flex;flex-direction:column}
#error{flex:none;max-height:30%;overflow:auto;background:#fff0ef;color:#922e25;margin:0;padding:16px 20px;white-space:pre-wrap;border-bottom:1px solid #e8bbb7}
#error[hidden]{display:none}
#stage{flex:1;min-height:0;overflow:auto;display:flex;padding:32px;overscroll-behavior:contain}
#slide{position:relative;flex:none;margin:auto;background:white;box-shadow:0 3px 24px #19212d20;overflow:hidden}
#slide{user-select:text;-webkit-user-select:text}
#empty{margin:auto;color:#737d8e}
footer{display:flex;align-items:center;gap:14px;padding:0 16px;background:#fff;border-top:1px solid #d4d9e2;font-size:12px}
#count{min-width:90px}footer .hint{flex:1;color:#737d8e}footer button{padding:3px 10px}footer select{padding:3px 8px}
#presentation-controls{display:none}
body.presenting{grid-template-rows:minmax(0,1fr);background:#111}
.presenting header,.presenting footer,.presenting .filmstrip,.presenting #chat{display:none}
.presenting .workspace{grid-template-columns:minmax(0,1fr)}.presenting #stage{padding:0;background:#111}.presenting #slide{box-shadow:none}
.presenting #presentation-controls{display:flex;position:fixed;bottom:16px;left:50%;transform:translateX(-50%);align-items:center;gap:12px;background:#202735e8;color:white;padding:6px;border-radius:8px;opacity:0;transition:opacity .15s}
.presenting #presentation-controls:hover,.presenting #presentation-controls:focus-within{opacity:1}
#presentation-controls button{background:transparent;color:white;border-color:#5d6575}
@media(max-width:700px){.workspace{grid-template-columns:140px minmax(0,1fr)}.filmstrip{padding:12px 5px}header{padding:0 12px;gap:10px}.badge,footer .hint{display:none}#stage{padding:16px}footer{gap:8px}#status{font-size:11px}.download{padding:7px 8px}}
#chat{min-height:0;display:flex;flex-direction:column;background:#fff;border-left:1px solid #d4d9e2}
.chat-heading{display:flex;align-items:center;gap:8px;padding:16px;border-bottom:1px solid #e6e9ef}.chat-heading strong{flex:1}.chat-heading button{padding:5px 8px}
#messages{flex:1;overflow:auto;padding:16px;display:flex;flex-direction:column;gap:16px;overscroll-behavior:contain}
.chat-message{white-space:pre-wrap;overflow-wrap:anywhere;line-height:1.6}.chat-message.user{background:#eef2ff;border-radius:8px;padding:10px 12px}.chat-message small{display:block;color:#697386;font-size:11px;margin-bottom:4px}.chat-intro{color:#737d8e;line-height:1.7}
#chat-form{padding:14px;border-top:1px solid #e6e9ef}#chat-context{font-size:11px;color:#4967dd;margin-bottom:8px}#chat-input{font:inherit;resize:vertical;min-height:90px;max-height:200px;width:100%;padding:10px;border:1px solid #d4d9e2;border-radius:8px} .chat-actions{display:flex;gap:8px;margin-top:8px;align-items:center}.chat-actions select{min-width:0;flex:1}#chat-send{background:#293c73;color:white}#chat-status{padding:0 14px 12px;color:#667085;font-size:11px;white-space:pre-wrap;max-height:90px;overflow:auto}
body.chat-hidden .workspace{grid-template-columns:190px minmax(0,1fr)}body.chat-hidden #chat{display:none}
body.presenting .workspace{grid-template-columns:minmax(0,1fr)}
@media(max-width:1000px){.workspace{grid-template-columns:120px minmax(0,1fr) 300px}.filmstrip{padding:12px 5px}#stage{padding:16px}}
@media(max-width:700px){.workspace{grid-template-columns:minmax(0,1fr);grid-template-rows:minmax(160px,1fr) minmax(200px,1fr)}.filmstrip{display:none}#chat{border-top:1px solid #d4d9e2}.chat-heading{padding:8px 12px}#chat-form{padding:8px 12px}#chat-input{min-height:50px}body.chat-hidden .workspace,.presenting .workspace{grid-template-columns:minmax(0,1fr);grid-template-rows:minmax(0,1fr)}.brand{font-size:14px}header{gap:6px}.download{font-size:11px}}
@media(prefers-reduced-motion:reduce){*{scroll-behavior:auto!important;transition:none!important}}
</style>
<header><span class="brand">Office Kit</span><span class="badge">Preview</span><span id="status" role="status">Building…</span><button id="toggle-chat" aria-expanded="true" aria-controls="chat">Chat</button><button id="present" disabled>Present</button><a class="download" href="/deck.pptx">Download PPTX</a></header>
<div class="workspace">
<nav class="filmstrip" aria-label="Slides"><h2>Slides</h2><ol id="thumbnails"></ol></nav>
<main aria-label="Slide viewer"><pre id="error" role="alert" hidden></pre><div id="stage" tabindex="-1"><div id="empty">Waiting for slides…</div><div id="slide" hidden></div></div></main>
<aside id="chat" aria-label="Slide chat">
<div class="chat-heading"><strong>Edit with AI</strong><button id="chat-reset">New chat</button></div>
<div id="messages" role="log" aria-label="Conversation"><p class="chat-intro">Describe what you want to change. The current slide is included automatically. You can also ask about other slides or the whole deck.<br><br>Uses your locally installed Claude Code or Codex. Sign in through its CLI first. Sending a message allows it to edit this project.</p></div>
<form id="chat-form"><div id="chat-context">No slide selected</div><label for="chat-input" class="chat-intro">What would you like to change?</label><textarea id="chat-input" maxlength="16000" placeholder="Make this headline shorter…" required></textarea><div class="chat-actions"><select id="chat-provider" aria-label="AI provider"><option value="claude">Claude Code</option><option value="codex">Codex</option></select><button id="chat-stop" type="button" hidden>Stop</button><button id="chat-send" type="submit">Send</button></div></form>
<div id="chat-status" role="status">Ready</div>
</aside>
</div>
<footer><span id="count" aria-live="polite">No slides</span><span class="hint">Changes appear automatically · Text can be selected and copied</span><button id="prev" aria-label="Previous slide" disabled>‹</button><button id="next" aria-label="Next slide" disabled>›</button><label for="zoom">Zoom</label><select id="zoom"><option value="fit">Fit</option><option value="0.5">50%</option><option value="0.75">75%</option><option value="1">100%</option><option value="1.25">125%</option><option value="1.5">150%</option><option value="2">200%</option></select></footer>
<div id="presentation-controls"><button id="present-prev" aria-label="Previous slide">‹</button><span id="present-count"></span><button id="present-next" aria-label="Next slide">›</button><button id="exit-present">Exit · Esc</button></div>
<script>
let state={slides:[],error:null,aspectRatio:16/9},index=0,urls=[],presenting=false;
let displayedSvg;
const byId=id=>document.getElementById(id);
const stage=byId('stage'),slide=byId('slide'),thumbnails=byId('thumbnails');
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
  byId('chat-context').textContent=state.slides.length?'Context: '+count:'Context: whole project';
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
let chatRequest=0,chatBusy=false,lastMessages='';
function renderChat(chat){
  chatBusy=chat.running;
  byId('chat-send').disabled=chat.running;
  byId('chat-reset').disabled=chat.running;
  byId('chat-provider').disabled=chat.running;
  byId('chat-stop').hidden=!chat.running;
  byId('chat-status').textContent=chat.status;
  if(chat.messages.length)byId('chat-provider').value=chat.provider;
  const serialized=JSON.stringify(chat.messages);
  if(serialized===lastMessages)return;
  lastMessages=serialized;
  const log=byId('messages'),atBottom=log.scrollHeight-log.scrollTop-log.clientHeight<60;
  log.replaceChildren();
  if(!chat.messages.length){const intro=document.createElement('p');intro.className='chat-intro';intro.textContent='Ask for a focused slide edit or a change across the deck. Uses your locally installed CLI and its login. Sending allows edits to this project.';log.append(intro);}
  for(const message of chat.messages){
    const item=document.createElement('div'),label=document.createElement('small'),body=document.createElement('div');
    item.className='chat-message '+message.role;
    label.textContent=message.role==='user'?'You · '+(message.context?.slide?'Slide '+message.context.slide:'Whole project'):'Assistant';
    body.textContent=message.text||(chat.running?'Working…':'No response');item.append(label,body);log.append(item);
  }
  if(atBottom)log.scrollTop=log.scrollHeight;
}
async function refreshChat(){
  const id=++chatRequest;
  try{const response=await fetch('/chat');if(!response.ok)throw new Error('Chat unavailable');const chat=await response.json();if(id===chatRequest)renderChat(chat);}
  catch(error){if(id===chatRequest)byId('chat-status').textContent=error.message;}
}
async function chatAction(path,payload){
  ++chatRequest;
  const response=await fetch(path,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
  const result=await response.json();if(!response.ok)throw new Error(result.error||'Chat request failed');
  await refreshChat();
}
byId('chat-form').onsubmit=async event=>{
  event.preventDefault();if(chatBusy)return;
  const input=byId('chat-input'),message=input.value.trim();if(!message)return;
  chatBusy=true;byId('chat-send').disabled=true;
  try{await chatAction('/chat',{message,provider:byId('chat-provider').value,slide:state.slides.length?index:null,revision:state.revision??0});if(input.value.trim()===message)input.value='';}
  catch(error){byId('chat-status').textContent=error.message;chatBusy=false;byId('chat-send').disabled=false;}
};
byId('chat-input').onkeydown=event=>{if(event.key==='Enter'&&(event.metaKey||event.ctrlKey)&&!event.isComposing){event.preventDefault();byId('chat-form').requestSubmit();}};
for(const action of ['stop','reset'])byId('chat-'+action).onclick=()=>chatAction('/chat/'+action,{}).catch(error=>{byId('chat-status').textContent=error.message;});
byId('toggle-chat').onclick=()=>{const hidden=document.body.classList.toggle('chat-hidden');byId('toggle-chat').setAttribute('aria-expanded',String(!hidden));resize();};
void refreshChat();
const events=new EventSource('/events');events.onmessage=event=>{if(event.data==='chat'){void refreshChat();return;}if(event.data==='ready'){delete state.revision;void refreshChat();}void refresh();};events.onerror=()=>{byId('status').textContent='Reconnecting…'};refresh();
</script></html>`;
