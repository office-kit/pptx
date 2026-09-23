import { previewI18nScript } from './preview-i18n.ts';
import { transitionScript } from './transition-script.ts';
import { previewStyles } from './styles.ts';
export const page = `<!doctype html>
<html lang="en">
<meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>Office Kit — PowerPoint preview</title>
<link rel="stylesheet" href="/terminal.css">
<style>
${previewStyles}
.slide-edit-tools{display:flex;gap:6px;padding:8px 14px;flex-shrink:0}.slide-edit-tools{align-items:center;flex-wrap:wrap}.slide-edit-tools [data-undo],.slide-edit-tools [data-redo]{width:32px;height:32px;padding:0;display:grid;place-items:center;border:1px solid transparent;border-radius:8px;background:transparent;color:#625873;box-shadow:none}.slide-edit-tools [data-undo]:hover:not(:disabled),.slide-edit-tools [data-redo]:hover:not(:disabled){background:#ede7fa;border-color:#d6cbea;color:#513399}.slide-edit-tools button:disabled{opacity:.35}.slide-edit-tools button:focus-visible{outline:2px solid #8966df;outline-offset:2px}.slide-edit-tools [hidden]{display:none}.slide-edit-tools span{font-size:11px;color:#625873;flex:1}.slide-edit-tools small{font-size:11px;color:#634396}.slide-hover{position:fixed;border:1px solid #b094ff;pointer-events:none;z-index:19;border-radius:2px}.slide-text-input{position:fixed;z-index:22;resize:none;background:#202438;color:#fff;border:2px solid #b094ff;border-radius:3px;padding:0;box-shadow:0 0 0 3px #b094ff30;line-height:1.15}.slide-text-input[hidden],.slide-hover[hidden]{display:none}.slide-selection{position:fixed;border:2px solid #b094ff;background:#a080ff18;pointer-events:none;z-index:20;box-shadow:0 0 0 1px #201632}.slide-edit-panel{position:fixed;z-index:21;width:min(380px,calc(100vw - 16px));padding:14px;background:#202438;border:1px solid #776591;border-radius:14px;box-shadow:0 16px 50px #0008;display:flex;flex-direction:column;gap:12px}.slide-edit-panel[hidden],.slide-selection[hidden]{display:none}.selection-heading,.selection-actions{display:flex;align-items:center;gap:8px}.selection-heading strong{flex:1}.slide-edit-panel textarea{width:100%;min-height:90px;max-height:35vh;resize:vertical;background:#141a2b;color:#f4efff;border:1px solid #514969;border-radius:8px;padding:10px;font:inherit}.selection-actions{flex-wrap:wrap}.selection-actions select{min-width:80px;flex:1}.slide-edit-panel small{color:#d1c9e3;line-height:1.5}.presenting .slide-edit-tools{display:none}

#agent-workspace{position:relative;flex:1;min-height:0;min-width:0;display:flex;overflow:auto}.agent-pane{position:absolute;padding:3px 5px;display:flex;flex-direction:column;flex:1;min-width:0;min-height:0;overflow:hidden}.agent-pane iframe{border:0;width:100%;flex:1;min-height:0;border-radius:0 0 9px 9px}.agent-tools{display:flex;gap:4px;align-items:center;flex-wrap:wrap;padding:6px;background:#20253a;border:1px solid #34334e;border-radius:9px 9px 0 0}.agent-tools span{flex:1;font-size:10px;letter-spacing:.04em;color:#b3abc9}.agent-tools svg{width:14px;height:14px}.agent-pane:focus-within .agent-tools{border-color:#8b6ed577;background:#29243e}.agent-tools button{font-size:11px;padding:5px;background:transparent;border:0;display:grid;place-items:center}.agent-divider{position:absolute;z-index:2;background:#101422;cursor:col-resize;touch-action:none}.agent-divider.vertical{cursor:row-resize}.agent-divider:hover,.agent-divider:focus-visible{background:#9b87ff}.splitting-agents iframe,.resizing-chat iframe{pointer-events:none}
.thumbnail[data-skipped="true"] .slide-number{text-decoration:line-through}
.thumbnail[data-skipped="true"] img{opacity:.6}
#editor-frame{display:none;border:0;width:100%;height:100%;flex:1;min-height:0}
body.editing:not(.presenting){grid-template-rows:60px minmax(0,1fr)}
.editing:not(.presenting) .workspace{--filmstrip-width:0px;grid-template-columns:minmax(0,1fr) clamp(280px,var(--chat-width),calc(100vw - 500px))}
.editing.chat-hidden:not(.presenting) .workspace{grid-template-columns:minmax(0,1fr)}
.editing:not(.presenting) .filmstrip,.editing:not(.presenting) footer,.editing:not(.presenting) #stage,.editing:not(.presenting) #present,.editing:not(.presenting) #presenter,.editing:not(.presenting) .download{display:none}
.editing:not(.presenting) #editor-frame{display:block}
@media(max-width:900px){.editing:not(.presenting) .workspace{grid-template-columns:minmax(0,1fr)}.editing:not(.presenting) #chat{display:none}}
</style>
<body class="editing"><header><span class="brand"><svg viewBox="0 0 36 36" aria-hidden="true"><defs><linearGradient id="brand-gradient" x2="1" y2="1"><stop stop-color="#c39aff"/><stop offset="1" stop-color="#7165f4"/></linearGradient></defs><rect x="2" y="2" width="32" height="32" rx="10" fill="url(#brand-gradient)"/><path d="M11 10h9l6 6v10H11z" fill="none" stroke="white" stroke-width="1.8" stroke-linejoin="round"/><path d="M20 10v7h6M15 21h7M15 25h4" fill="none" stroke="white" stroke-width="1.8" stroke-linecap="round"/><path d="m28 3 1.2 3.8L33 8l-3.8 1.2L28 13l-1.2-3.8L23 8l3.8-1.2z" fill="#e7dbff"/></svg>Office <em>Kit</em></span><span class="badge">Studio</span><span id="status" role="status">Building…</span><button id="toggle-editor" aria-pressed="true">Preview</button><button id="toggle-chat" aria-expanded="true" aria-controls="chat">✦ Agents</button><button id="present" disabled>Present</button><button id="presenter" disabled>Presenter view</button><a class="download" href="/deck.pptx">Download PPTX</a></header>
<div class="workspace">
<nav class="filmstrip" aria-label="Slides"><h2>Slides</h2><ol id="thumbnails"></ol></nav>
<main aria-label="Slide viewer"><iframe id="editor-frame" title="Presentation editor"></iframe><pre id="error" role="alert" hidden></pre><div id="stage" tabindex="-1"><div id="empty">Waiting for slides…</div><div id="slide" hidden></div></div></main>
<aside id="chat" aria-label="Slide chat">
<div id="chat-resizer" role="separator" tabindex="0" aria-label="Chat width" aria-orientation="vertical" aria-controls="chat" title="Drag to resize · Double-click to reset"></div>
<div class="workspace-heading"><span>✦ AI WORKSPACE</span><b>LOCAL</b></div><div id="agent-workspace"></div><div id="chat-context" hidden></div></aside>
</div>
<footer><span id="count" aria-live="polite">No slides</span><span class="hint">Select an area to ask AI · Edit text directly</span><button id="prev" aria-label="Previous slide" disabled>‹</button><button id="next" aria-label="Next slide" disabled>›</button><label for="zoom">Zoom</label><select id="zoom"><option value="fit">Fit</option><option value="0.5">50%</option><option value="0.75">75%</option><option value="1">100%</option><option value="1.25">125%</option><option value="1.5">150%</option><option value="2">200%</option></select></footer>
<div id="presentation-controls"><button id="present-prev" aria-label="Previous slide">‹</button><span id="present-count"></span><span id="present-note" role="status" hidden></span><button id="animation-retry" hidden></button><button id="present-next" aria-label="Next slide">›</button><button id="exit-present">Exit · Esc</button></div>
<script>
let state={slides:[],error:null,aspectRatio:16/9},index=0,urls=[],presenting=false;
let displayedSvg;
let presenterWindow;
${previewI18nScript}
function findSlide(start,step,skipHidden=presenting){
 for(let i=start;i>=0&&i<state.slides.length;i+=step)if(!skipHidden||!state.hiddenSlides?.[i])return i;
 return -1;
}
function moveSlide(step,skipHidden=presenting,focusThumbnail=false){
 // The slide has effects and the player is still on its way: moving on now
 // would skip them, so the click waits rather than costing the viewer content.
 if(animationsWaiting())return;
 // A build comes before the slide: the click that would move on first plays
 // whatever the current slide still has to show, in either direction.
 if(animationPlayer&&(step>0?animationPlayer.advance():animationPlayer.back()))return;
 const next=findSlide(index+step,step,skipHidden);
 // Arriving backwards lands on a slide that has already played out, the way
 // PowerPoint shows it; arriving forwards starts its build from the top.
 if(next>=0)selectSlide(next,focusThumbnail,true,step<0?'end':'start');
}
let advanceTimer,advanceKey;
function scheduleAdvance(){
  const delay=state.transitions?.[index]?.advanceAfterMs;
  // A slide that still has effects to play — or one whose last click is still
  // playing out — is not finished, whatever its transition says about
  // advancing itself.
  const settled=!animationsWaiting()&&(!animationPlayer||!(animationPlayer.pending||animationPlayer.running));
  const key=presenting&&settled&&findSlide(index+1,1)>=0&&Number.isFinite(delay)&&delay>=0?index+':'+delay:null;
  if(key===advanceKey)return;
  clearTimeout(advanceTimer);advanceKey=key;
  if(key===null)return;
  // OOXML unsigned milliseconds can exceed the browser's signed timer limit.
  const deadline=performance.now()+delay;
  function tick(){
    const remaining=deadline-performance.now();
    if(remaining>0)advanceTimer=setTimeout(tick,Math.min(remaining,2147483647));
    else {advanceKey=null;moveSlide(1);}
  }
  advanceTimer=setTimeout(tick,Math.min(delay,2147483647));
}
const byId=id=>document.getElementById(id);
// Next and previous move through the slide's own build before they move
// through the deck, so the last slide's remaining effects are still reachable
// and the first step back is the one inside this slide.
// True while this slide has effects but the player has not arrived yet. Its
// clicks belong to the build, so they are held rather than spent on the deck.
function animationsWaiting(){return animationLoad==='loading'&&presenting&&animationStepsAt(index).length>0;}
function animationsPending(){return animationsWaiting()||Boolean(animationPlayer&&animationPlayer.pending);}
function animationsBehind(){return Boolean(animationPlayer&&animationPlayer.cursor>0);}
function updateNavigationButtons(){
  for(const id of ['prev','present-prev'])byId(id).disabled=findSlide(index-1,-1)<0&&!animationsBehind();
  for(const id of ['next','present-next'])byId(id).disabled=findSlide(index+1,1)<0&&!animationsPending();
  byId('present').disabled=findSlide(0,1,true)<0;
  byId('presenter').disabled=byId('present').disabled;
}
const stage=byId('stage'),slide=byId('slide'),thumbnails=byId('thumbnails');
const workspace=document.querySelector('.workspace'),chatResizer=byId('chat-resizer');
const chatWidthKey='office-kit-chat-width',minChatWidth=280,minStageWidth=240;
function chatWidthLimits(){return {min:minChatWidth,max:Math.max(minChatWidth,workspace.clientWidth-document.querySelector('.filmstrip').getBoundingClientRect().width-minStageWidth)};}
function updateChatWidthAria(){
  const {min,max}=chatWidthLimits();
  chatResizer.setAttribute('aria-valuemin',String(min));chatResizer.setAttribute('aria-valuemax',String(Math.round(max)));
  const width=Math.round(byId('chat').getBoundingClientRect().width);
  chatResizer.setAttribute('aria-valuenow',String(width));chatResizer.setAttribute('aria-valuetext',width+(previewLocale==='ja'?' ピクセル':' pixels'));
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
${transitionScript}
// The slide's object animations. One player per slide, kept while the slide and
// its timing stay as they are: an unrelated rebuild (new speaker notes, say)
// must not throw away how far the viewer has clicked through a build.
//
// The player is the editor's module, served as one build, so the show, the
// presenter window and the editor panel cannot disagree about a timing tree.
let animationPlayer,animationPlayerKey,createAnimationPlayer,pendingAnimationPosition;
// 'loading' | 'ready' | 'failed'. A slide show holds its clicks while the
// player is on its way; if it never arrives, the viewer is told so and can ask
// for it again rather than watching a deck quietly skip its animations.
let animationLoad='loading',animationLoadAttempt=0;
function loadAnimationPlayer(){
  if(animationLoad==='ready')return;
  animationLoad='loading';updateAnimationNotice();
  // A module that failed to load stays failed in the browser's module map, so
  // asking again means asking for a different URL.
  const url='/animation-player.js'+(animationLoadAttempt++?'?retry='+animationLoadAttempt:'');
  import(url).then(module=>{
    createAnimationPlayer=module.createAnimationPlayer;animationLoad='ready';
    const position=pendingAnimationPosition;pendingAnimationPosition=undefined;
    if(position!==undefined)syncAnimationPlayer(position);
    updateNavigationButtons();updateAnimationNotice();scheduleAdvance();
  },error=>{
    animationLoad='failed';console.warn('Could not load the animation player',error);
    updateNavigationButtons();updateAnimationNotice();scheduleAdvance();
  });
}
const animationStepsAt=i=>state.animations?.[i]??[];
const animationKeyAt=i=>i+'\u0000'+(state.slides[i]??'')+'\u0000'+JSON.stringify(animationStepsAt(i));
// A transition draws the arriving slide into a layer of its own, beside a layer
// holding the slide being left — whose shapes carry ids of their own, from a
// different slide's id space. Only the arriving one is this slide.
const slideRoot=()=>canvas.querySelector('.transition-layer:not(.transition-old)')??canvas;
function updateAnimationNotice(){
  const note=byId('present-note'),retry=byId('animation-retry');
  const slideHasEffects=presenting&&animationStepsAt(index).length>0;
  retry.hidden=!(slideHasEffects&&animationLoad==='failed');
  retry.textContent=previewLocale==='ja'?'再試行':'Try again';
  if(slideHasEffects&&animationLoad==='failed'){
    note.hidden=false;
    note.textContent=previewLocale==='ja'
      ?'アニメーションを読み込めませんでした。このスライドの効果は再生されません。'
      :'The animations could not be loaded, so this slide’s effects are not played.';
    return;
  }
  if(animationsWaiting()){
    note.hidden=false;
    note.textContent=previewLocale==='ja'?'アニメーションを準備中…':'Preparing animations…';
    return;
  }
  const skipped=animationPlayer?animationPlayer.unsupported:[];
  note.hidden=skipped.length===0;
  note.textContent=skipped.length===0?'':previewLocale==='ja'
    ?'このスライドの '+skipped.length+' 件のアニメーションは未対応のため再生されません'
    :skipped.length===1
      ?'1 animation on this slide is not supported yet, so it is not played'
      :skipped.length+' animations on this slide are not supported yet, so they are not played';
}
function syncAnimationPlayer(position){
  // Animations belong to the slide show. The still preview shows the slide as
  // it ends, which is what the renderer drew.
  const wanted=presenting&&animationStepsAt(index).length>0;
  if(!wanted){
    if(animationPlayer){animationPlayer.dispose();animationPlayer=undefined;animationPlayerKey=undefined;}
    updateAnimationNotice();
    return;
  }
  if(!createAnimationPlayer){
    // Still on its way, or it never arrived. The slide keeps the picture the
    // renderer drew; this position is applied the moment the player is there.
    pendingAnimationPosition=position;
    updateAnimationNotice();
    return;
  }
  const key=animationKeyAt(index);
  if(!animationPlayer||animationPlayerKey!==key){
    // A changed timing makes the old cursor meaningless — the effects it
    // counted are not the effects there are now.
    animationPlayer?.dispose();
    animationPlayer=createAnimationPlayer({
      root:slideRoot,
      steps:animationStepsAt(index),
      reducedMotion:()=>reducedMotion.matches,
      onChange:()=>{updateNavigationButtons();updatePresenter();scheduleAdvance();},
    });
    animationPlayerKey=key;
    position==='end'?animationPlayer.finish():animationPlayer.reset();
    updateAnimationNotice();
    return;
  }
  if(position==='end')animationPlayer.finish();
  else if(position==='start')animationPlayer.reset();
  // 'keep' leaves the cursor where the viewer left it.
  updateAnimationNotice();
}
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
function selectSlide(next,focusThumbnail=false,reveal=true,position='start'){
  const previousIndex=index;
  index=Math.max(0,Math.min(next,state.slides.length-1));
  if(presenting&&state.hiddenSlides?.[index]){
    const forward=findSlide(index,1);index=forward>=0?forward:Math.max(0,findSlide(index,-1));
  }
  const count=slideCount();
  byId('chat-context').textContent=state.slides.length?count:pt('Whole project');
  byId('chat-context').dataset.focus=JSON.stringify({slide:state.slides.length?index:null,revision:state.revision??0});
  window.dispatchEvent(new Event('agent-focus'));
  if(document.body.classList.contains('editing')&&editorFocus)applyEditorFocus();
  byId('count').textContent=count;byId('present-count').textContent=count;
  byId('zoom').disabled=!state.slides.length;
  slide.hidden=!state.slides.length;byId('empty').hidden=!!state.slides.length;
  const svg=state.slides[index];
  if(svg!==displayedSvg||previousIndex!==index){
    renderSlide(svg,presenting&&previousIndex!==index?state.transitions?.[index]:null);
  }
  syncAnimationPlayer(position);
  updateNavigationButtons();
  slide.setAttribute('aria-label',slideLabel(index));
  for(const [position,item] of Array.from(thumbnails.children).entries()){
    const button=item.firstElementChild,selected=position===index;
    button.setAttribute('aria-current',String(selected));button.tabIndex=selected?0:-1;
    button.dataset.skipped=String(!!state.hiddenSlides?.[position]);
    button.title=state.hiddenSlides?.[position]?pt('Skipped during presentation'):'';
    if(selected){if(reveal)button.scrollIntoView({block:'nearest'});if(focusThumbnail)button.focus({preventScroll:true});}
  }
  resize();
  if(!transitionCleanup)scheduleAdvance();
  updatePresenter();
}
function update(updated){
  const focusedThumbnail=thumbnails.contains(document.activeElement);
  const previous=state;
  state=updated;
  connectionLost=false;updatePreviewStatus();
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
      button.className='thumbnail';button.setAttribute('aria-label',slideLabel(i));button.onclick=()=>selectSlide(i,true);
      number.className='slide-number';number.textContent=String(i+1);
      image.alt='';image.draggable=false;image.loading='lazy';
      button.append(number,image);item.append(button);thumbnails.append(item);
    }
    item.querySelector('img').src=urls[i];
    if(oldUrl)URL.revokeObjectURL(oldUrl);
  });
  if(presenting&&findSlide(0,1,true)<0)void exitPresentation();
  selectSlide(index,focusedThumbnail,false,'keep');
}
function setPresenting(value){
  cancelTransition();
  presenting=value&&findSlide(0,1,true)>=0;document.body.classList.toggle('presenting',presenting);selectSlide(index);
  if(presenting)stage.focus();else{
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
  catch{byId('exit-present').textContent=pt('Exit view · Esc');}
};
function updatePresenter(){
 if(!presenterWindow||presenterWindow.closed)return;
 // The presenter runs the same player over its own copy of the slide, from the
 // same cursor and the same moment within the stop in hand, so what it shows
 // is what the audience sees — including an effect still fading in.
 const progress=animationPlayer?animationPlayer.progress:null;
 const animationSteps=animationPlayer?animationStepsAt(index):null;
 presenterWindow.postMessage({type:'presenter-state',index,count:state.slides.length,animationSteps,current:state.slides[index]??null,next:state.slides[findSlide(index+1,1,true)]??null,hasPrevious:findSlide(index-1,-1,true)>=0||animationsBehind(),hasNext:findSlide(index+1,1,true)>=0||animationsPending(),animation:progress,notes:state.notes?.[index]??'',aspectRatio:state.aspectRatio,locale:previewLocale,presenting},location.origin);
}
byId('presenter').onclick=()=>{
 if(presenterWindow&&!presenterWindow.closed){setPresenting(true);presenterWindow.focus();return;}
 presenterWindow=window.open('/presenter','office-kit-presenter','popup,width=1100,height=800');
 if(presenterWindow)setPresenting(true);
 else byId('status').textContent=pt('Allow popups to open presenter view.');
};
window.addEventListener('message',event=>{
 if(event.origin!==location.origin||event.source!==presenterWindow||event.data?.type!=='presenter-command')return;
 if(event.data.action==='ready')updatePresenter();
 else if(event.data.action==='next')moveSlide(1,true);
 else if(event.data.action==='previous')moveSlide(-1,true);
 else if(event.data.action==='exit')void exitPresentation();
 else if(event.data.action==='jump'&&Number.isInteger(event.data.index)&&event.data.index>=0&&event.data.index<state.slides.length)selectSlide(event.data.index);
});
byId('exit-present').onclick=exitPresentation;
byId('animation-retry').onclick=()=>{loadAnimationPlayer();};
loadAnimationPlayer();
document.addEventListener('fullscreenchange',()=>{if(!document.fullscreenElement&&presenting)setPresenting(false);});
for(const id of ['prev','present-prev'])byId(id).onclick=()=>moveSlide(-1);
for(const id of ['next','present-next'])byId(id).onclick=()=>moveSlide(1);
byId('zoom').onchange=resize;
stage.onclick=event=>{
  const link=event.composedPath().find(node=>node instanceof Element&&node.localName==='a');
  if(link){
    const href=link.getAttribute('href')??link.getAttributeNS('http://www.w3.org/1999/xlink','href')??'';
    if(href.startsWith('#slide-')){
      event.preventDefault();
      const number=Number(href.slice(7));
      if(Number.isInteger(number)&&number>=1&&number<=state.slides.length)selectSlide(number-1);
    }
    return;
  }
  if(!presenting||getSelection().toString())return;
  // 'advClick' says whether a click moves to the next *slide*. A build still
  // belongs to this one, so its remaining effects play either way.
  if(animationsPending()||state.transitions?.[index]?.advanceOnClick!==false)moveSlide(1);
};
document.addEventListener('keydown',event=>{
  if(event.key==='Escape'&&presenting){event.preventDefault();void exitPresentation();return;}
  if(event.altKey||event.ctrlKey||event.metaKey||event.target.closest('select,input,textarea,[contenteditable]'))return;
  const focusThumbnail=thumbnails.contains(document.activeElement);
  let step=0,jump=-1;
  if(['ArrowLeft','ArrowUp','PageUp'].includes(event.key))step=-1;
  else if(['ArrowRight','ArrowDown','PageDown'].includes(event.key))step=1;
  else if(event.key==='Home')jump=findSlide(0,1);
  else if(event.key==='End')jump=findSlide(state.slides.length-1,-1);
  else if(event.key===' '&&presenting&&!event.target.closest('button,a'))step=event.shiftKey?-1:1;
  else return;
  event.preventDefault();
  // Home and End are a jump, not a step: they land on a slide with its build
  // at the start rather than walking through the one in hand.
  if(step!==0)moveSlide(step,presenting,focusThumbnail);
  else if(jump>=0)selectSlide(jump,focusThumbnail);
});
new ResizeObserver(resize).observe(stage);
let refreshId=0;
async function refresh(){
  const id=++refreshId;
  try{const response=await fetch('/state'+(state.revision===undefined?'':'?since='+state.revision));if(!response.ok)throw new Error('Preview unavailable');const updated=await response.json();if(id===refreshId){if(updated.changes){updated.slides=state.slides.slice(0,updated.count);updated.slides.length=updated.count;for(const [position,svg] of Object.entries(updated.changes))updated.slides[Number(position)]=svg;}update(updated);}}
  catch{if(id===refreshId){connectionLost=true;updatePreviewStatus();}}
}
byId('toggle-chat').onclick=()=>{const hidden=document.body.classList.toggle('chat-hidden');byId('toggle-chat').setAttribute('aria-expanded',String(!hidden));resize();};
const editorFrame=byId('editor-frame');
let editorFocus;
function setEditorMode(editing){
 document.body.classList.toggle('editing',editing);
 if(editing&&!editorFrame.getAttribute('src'))editorFrame.src='/editor';
 byId('toggle-editor').setAttribute('aria-pressed',String(editing));
 byId('toggle-editor').textContent=pt(editing?'Preview':'Edit');
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
 if(previewLocale!==editorFocus.locale&&(editorFocus.locale==='ja'||editorFocus.locale==='en')){previewLocale=editorFocus.locale;updatePreviewLabels();}
 updatePresenter();
 if(document.body.classList.contains('editing'))applyEditorFocus();
 byId('toggle-editor').textContent=pt(document.body.classList.contains('editing')?'Preview':'Edit');
});
updatePreviewLabels();
const events=new EventSource('/events');events.onmessage=event=>{if(event.data==='history'){window.dispatchEvent(new Event('agent-history'));return;}if(event.data==='chat'){window.dispatchEvent(new Event('agent-chat'));return;}if(event.data==='ready'){delete state.revision;}void refresh();};events.onerror=()=>{connectionLost=true;updatePreviewStatus()};refresh();
</script><script type="module" src="/terminal.js"></script></html>`;
