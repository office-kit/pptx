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
.workspace{display:grid;grid-template-columns:224px minmax(0,1fr);min-height:0}
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
#slide{flex:none;margin:auto;background:white;box-shadow:0 3px 24px #19212d20;overflow:hidden}
#slide iframe{display:block;width:100%;height:100%;border:0;pointer-events:none}
#empty{margin:auto;color:#737d8e}
footer{display:flex;align-items:center;gap:14px;padding:0 16px;background:#fff;border-top:1px solid #d4d9e2;font-size:12px}
#count{min-width:90px}footer .hint{flex:1;color:#737d8e}footer button{padding:3px 10px}footer select{padding:3px 8px}
#presentation-controls{display:none}
body.presenting{grid-template-rows:minmax(0,1fr);background:#111}
.presenting header,.presenting footer,.presenting .filmstrip{display:none}
.presenting .workspace{grid-template-columns:minmax(0,1fr)}.presenting #stage{padding:0;background:#111}.presenting #slide{box-shadow:none}
.presenting #presentation-controls{display:flex;position:fixed;bottom:16px;left:50%;transform:translateX(-50%);align-items:center;gap:12px;background:#202735e8;color:white;padding:6px;border-radius:8px;opacity:0;transition:opacity .15s}
.presenting #presentation-controls:hover,.presenting #presentation-controls:focus-within{opacity:1}
#presentation-controls button{background:transparent;color:white;border-color:#5d6575}
@media(max-width:700px){.workspace{grid-template-columns:140px minmax(0,1fr)}.filmstrip{padding:12px 5px}header{padding:0 12px;gap:10px}.badge,footer .hint{display:none}#stage{padding:16px}footer{gap:8px}#status{font-size:11px}.download{padding:7px 8px}}
@media(prefers-reduced-motion:reduce){*{scroll-behavior:auto!important;transition:none!important}}
</style>
<header><span class="brand">Office Kit</span><span class="badge">Preview</span><span id="status" role="status">Building…</span><button id="present" disabled>Present</button><a class="download" href="/deck.pptx">Download PPTX</a></header>
<div class="workspace">
<nav class="filmstrip" aria-label="Slides"><h2>Slides</h2><ol id="thumbnails"></ol></nav>
<main aria-label="Slide viewer"><pre id="error" role="alert" hidden></pre><div id="stage" tabindex="-1"><div id="empty">Waiting for slides…</div><div id="slide" hidden></div></div></main>
</div>
<footer><span id="count" aria-live="polite">No slides</span><span class="hint">Changes appear automatically · View only</span><button id="prev" aria-label="Previous slide" disabled>‹</button><button id="next" aria-label="Next slide" disabled>›</button><label for="zoom">Zoom</label><select id="zoom"><option value="fit">Fit</option><option value="0.5">50%</option><option value="0.75">75%</option><option value="1">100%</option><option value="1.25">125%</option><option value="1.5">150%</option><option value="2">200%</option></select></footer>
<div id="presentation-controls"><button id="present-prev" aria-label="Previous slide">‹</button><span id="present-count"></span><button id="present-next" aria-label="Next slide">›</button><button id="exit-present">Exit · Esc</button></div>
<script>
let state={slides:[],error:null,aspectRatio:16/9},index=0,urls=[],presenting=false;
const byId=id=>document.getElementById(id);
const stage=byId('stage'),slide=byId('slide'),thumbnails=byId('thumbnails');
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
function selectSlide(next,focusThumbnail=false){
  index=Math.max(0,Math.min(next,state.slides.length-1));
  const count=state.slides.length?'Slide '+(index+1)+' of '+state.slides.length:'No slides';
  byId('count').textContent=count;byId('present-count').textContent=count;
  for(const id of ['prev','present-prev'])byId(id).disabled=index===0;
  for(const id of ['next','present-next'])byId(id).disabled=index>=state.slides.length-1;
  byId('present').disabled=!state.slides.length;
  byId('zoom').disabled=!state.slides.length;
  slide.hidden=!state.slides.length;byId('empty').hidden=!!state.slides.length;
  slide.replaceChildren();
  if(state.slides[index]){
    const frame=document.createElement('iframe');frame.sandbox='';frame.tabIndex=-1;frame.title='Slide '+(index+1);
    frame.srcdoc='<style>html,body{margin:0;width:100%;height:100%;overflow:hidden}svg{display:block;width:100%;height:100%}</style>'+state.slides[index];
    slide.append(frame);
  }
  for(const [position,item] of Array.from(thumbnails.children).entries()){
    const button=item.firstElementChild,selected=position===index;
    button.setAttribute('aria-current',String(selected));button.tabIndex=selected?0:-1;
    if(selected){button.scrollIntoView({block:'nearest'});if(focusThumbnail)button.focus({preventScroll:true});}
  }
  resize();
}
function update(updated){
  const focusedThumbnail=thumbnails.contains(document.activeElement);
  const changed=updated.slides.length!==state.slides.length||updated.slides.some((svg,i)=>svg!==state.slides[i]);
  state=updated;
  byId('status').textContent=state.error?'Build failed · showing last successful output':state.slides.length+' slides · Live';
  byId('error').textContent=state.error||'';byId('error').hidden=!state.error;
  document.documentElement.style.setProperty('--slide-ratio',String(state.aspectRatio));
  if(changed){
    for(const url of urls)URL.revokeObjectURL(url);
    urls=state.slides.map(svg=>URL.createObjectURL(new Blob([svg],{type:'image/svg+xml'})));
    thumbnails.replaceChildren(...urls.map((url,i)=>{
      const item=document.createElement('li'),button=document.createElement('button'),number=document.createElement('span'),image=document.createElement('img');
      button.className='thumbnail';button.setAttribute('aria-label','Slide '+(i+1));button.onclick=()=>selectSlide(i,true);
      number.className='slide-number';number.textContent=String(i+1);
      image.src=url;image.alt='';image.draggable=false;image.loading='lazy';
      button.append(number,image);item.append(button);return item;
    }));
    selectSlide(index,focusedThumbnail);
  }else resize();
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
stage.onclick=()=>{if(presenting)selectSlide(index+1);};
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
  try{const response=await fetch('/state');if(!response.ok)throw new Error('Preview unavailable');const updated=await response.json();if(id===refreshId)update(updated);}
  catch{if(id===refreshId)byId('status').textContent='Reconnecting…';}
}
const events=new EventSource('/events');events.onmessage=refresh;events.onerror=()=>{byId('status').textContent='Reconnecting…'};refresh();
</script></html>`;
