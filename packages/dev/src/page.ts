import { previewI18nScript } from './preview-i18n.ts';
import { transitionScript } from './transition-script.ts';
import { previewStyles } from './styles.ts';
export const page = `<!doctype html>
<html lang="en">
<meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>Presentation preview</title>
<link rel="stylesheet" href="/terminal.css">
<style>
${previewStyles}
.slide-edit-tools{display:flex;align-items:center;flex-wrap:wrap;gap:4px;padding:4px 12px;flex-shrink:0;background:var(--ok-ribbon);border-bottom:1px solid var(--ok-border)}.slide-edit-tools [data-undo],.slide-edit-tools [data-redo]{width:26px;height:26px;padding:0;display:grid;place-items:center;color:var(--ok-text-2)}.slide-edit-tools [data-undo]:hover:not(:disabled),.slide-edit-tools [data-redo]:hover:not(:disabled){color:var(--ok-text)}.slide-edit-tools span{font-size:11px;color:var(--ok-text-2);flex:1}.slide-edit-tools small{font-size:11px;color:var(--ok-accent)}.slide-hover{position:fixed;border:1px solid var(--ok-selected-border);pointer-events:none;z-index:19;border-radius:2px}.slide-text-input{position:fixed;z-index:22;resize:none;background:var(--ok-panel);color:var(--ok-text);border:2px solid var(--ok-selected-border);border-radius:var(--ok-radius);padding:0;line-height:1.15}.slide-selection{position:fixed;border:2px solid var(--ok-selected-border);background:color-mix(in srgb,var(--ok-selected-border) 10%,transparent);pointer-events:none;z-index:20}.slide-edit-panel{position:fixed;z-index:21;width:min(380px,calc(100vw - 16px));padding:12px;background:var(--ok-panel);color:var(--ok-text);border:1px solid var(--ok-border);border-radius:var(--ok-radius-lg);box-shadow:var(--ok-shadow-lg);display:flex;flex-direction:column;gap:10px}.selection-heading,.selection-actions{display:flex;align-items:center;gap:6px}.selection-heading strong{flex:1}.slide-edit-panel textarea{width:100%;min-height:90px;max-height:35vh;resize:vertical;background:var(--ok-panel);color:var(--ok-text);border:1px solid var(--ok-border-strong);border-radius:var(--ok-radius);padding:6px 8px}.slide-edit-panel textarea:focus{outline:2px solid var(--ok-selected-border);outline-offset:-1px}.selection-actions{flex-wrap:wrap}.selection-actions select{min-width:80px;flex:1}.selection-actions button{border-color:var(--ok-border-strong)}.slide-edit-panel small{color:var(--ok-text-2);line-height:1.5}.presenting .slide-edit-tools{display:none}
.editing:not(.presenting) .slide-edit-tools{display:none}

#agent-workspace{position:relative;flex:1;min-height:0;min-width:0;display:flex;overflow:auto}.agent-pane{position:absolute;display:flex;flex-direction:column;flex:1;min-width:0;min-height:0;overflow:hidden}.agent-pane iframe{border:0;width:100%;flex:1;min-height:0;background:var(--ok-panel)}.agent-tools{display:flex;gap:2px;align-items:center;padding:3px 6px 3px 12px;min-height:30px;background:var(--ok-panel-2);border-bottom:1px solid var(--ok-border)}.agent-tools span{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:11px;font-weight:600;color:var(--ok-text-2)}.agent-pane:focus-within .agent-tools span{color:var(--ok-text)}.agent-tools button{width:24px;height:24px;padding:0;display:grid;place-items:center;color:var(--ok-text-2)}.agent-tools button:hover:not(:disabled){color:var(--ok-text)}.agent-tools svg{width:14px;height:14px}.agent-divider{position:absolute;z-index:2;background:transparent;cursor:col-resize;touch-action:none}.agent-divider.vertical{cursor:row-resize}.agent-divider::after{content:"";position:absolute;background:var(--ok-border)}.agent-divider.horizontal::after{top:0;bottom:0;left:50%;width:1px}.agent-divider.vertical::after{left:0;right:0;top:50%;height:1px}.agent-divider:hover::after,.agent-divider:focus-visible::after{background:var(--ok-selected-border)}.agent-divider.horizontal:hover::after,.agent-divider.horizontal:focus-visible::after{width:2px}.agent-divider.vertical:hover::after,.agent-divider.vertical:focus-visible::after{height:2px}.agent-divider:focus-visible{outline:none}.splitting-agents iframe,.resizing-chat iframe{pointer-events:none}
#editor-frame{display:none;border:0;width:100%;height:100%;flex:1;min-height:0}
body.editing:not(.presenting){grid-template-rows:minmax(0,1fr)}
.editing:not(.presenting) .workspace{--filmstrip-width:0px;grid-template-columns:minmax(0,1fr) clamp(280px,var(--chat-width),calc(100vw - 500px))}
.editing.chat-hidden:not(.presenting) .workspace{grid-template-columns:minmax(0,1fr)}
.editing:not(.presenting) .filmstrip,.editing:not(.presenting) footer,.editing:not(.presenting) #stage{display:none}
.editing:not(.presenting) #editor-frame{display:block}
/* Too narrow to dock beside the editor: the pane floats over its right edge. */
@media(max-width:900px){.editing:not(.presenting) .workspace{grid-template-columns:minmax(0,1fr)}.editing:not(.presenting) #chat{position:fixed;z-index:30;top:0;right:0;bottom:0;width:min(380px,100vw);box-shadow:var(--ok-shadow-lg)}.editing:not(.presenting) #chat-resizer{display:none}}
</style>
<body class="editing chat-hidden">
<div class="workspace">
<nav class="filmstrip" aria-label="Slides"><h2>Slides</h2><ol id="thumbnails"></ol></nav>
<main aria-label="Slide viewer"><iframe id="editor-frame" title="Presentation editor"></iframe><pre id="error" role="alert" hidden></pre><div id="stage" tabindex="-1"><div id="empty">Waiting for slides…</div><div id="slide" hidden></div></div></main>
<aside id="chat" aria-labelledby="agents-title">
<div id="chat-resizer" role="separator" tabindex="0" aria-label="Chat width" aria-orientation="vertical" aria-controls="chat" title="Drag to resize · Double-click to reset"></div>
<div class="pane-head"><strong id="agents-title">Agents</strong><button id="close-chat" class="pane-close" aria-label="Close Agents" title="Close Agents">×</button></div><div id="agent-workspace"></div><div id="chat-context" hidden></div></aside>
</div>
<footer><span id="count" aria-live="polite">No slides</span><span id="status" role="status">Building…</span><span class="hint">Select an area to ask AI · Edit text directly</span><button id="prev" aria-label="Previous slide" disabled>‹</button><button id="next" aria-label="Next slide" disabled>›</button><button id="present" disabled>Present</button><button id="presenter" disabled>Presenter view</button><a class="download" href="/deck.pptx">Download PPTX</a><button id="toggle-chat" aria-pressed="false" aria-controls="chat"><svg viewBox="0 0 16 16" aria-hidden="true"><path d="M8 1.5 9.4 6.6 14.5 8 9.4 9.4 8 14.5 6.6 9.4 1.5 8 6.6 6.6Z" stroke-linejoin="round"/></svg><span>Agents</span></button><span class="status-views" role="group" aria-label="Presentation views"><button id="toggle-editor" aria-label="Normal" title="Normal (Esc)"><svg width="16" height="12" viewBox="0 0 16 12" aria-hidden="true"><rect x=".5" y=".5" width="15" height="11" rx="1"/><path d="M5 1v10"/></svg></button></span><label for="zoom">Zoom</label><select id="zoom"><option value="fit">Fit</option><option value="0.5">50%</option><option value="0.75">75%</option><option value="1">100%</option><option value="1.25">125%</option><option value="1.5">150%</option><option value="2">200%</option></select></footer><div id="notice" role="alert" hidden></div>
<div id="presentation-scrollbar" role="scrollbar" aria-label="Slide position" aria-controls="slide" aria-orientation="vertical" aria-valuemin="0" aria-valuemax="0" aria-valuenow="0" tabindex="0" hidden><span></span></div><div id="presentation-controls"><button id="present-prev" aria-label="Previous slide">‹</button><span id="present-count"></span><span id="rehearsal-timer" role="timer" aria-label="Rehearsal time" hidden></span><span id="present-note" role="status" hidden></span><button id="animation-retry" hidden></button><button id="present-next" aria-label="Next slide">›</button><button id="exit-present">Exit · Esc</button></div>
<script>
let state={slides:[],error:null,aspectRatio:16/9,showProperties:null,customShows:[]},index=0,urls=[],presenting=false;
let showOrder=[],showCursor=0,lastViewed=null,linkedShowId=null;
let showReturns=[];
// The fullscreen request of the show now starting. Starting or ending any show
// (Present, presenter view, Escape) drops it, so a grant arriving after that
// belongs to no show and is released instead of adopted.
let presentationFullscreen=false,fullscreenRequest=null;
function loopShow(){return showReturns.length===0&&(state.showProperties?.loop||state.showProperties?.mode?.kind==='kiosk');}
let displayedSvg;
let presenterWindow;
${previewI18nScript}
function rebuildShowOrder(){
 const previousSlide=showOrder[showCursor];
 const settings=state.showProperties;
 let order;
 if(linkedShowId!==null){
  order=state.customShows?.find(show=>show.id===linkedShowId)?.slideIndices??[];
 }else if(settings?.slides?.kind==='range'){
  const start=Math.max(1,settings.slides.start),end=Math.min(state.slides.length,settings.slides.end);
  order=Array.from({length:Math.max(0,end-start+1)},(_,offset)=>start-1+offset);
 }else if(settings?.slides?.kind==='customShow'){
  order=state.customShows?.find(show=>show.id===settings.slides.id)?.slideIndices??[];
 }else order=Array.from({length:state.slides.length},(_,slide)=>slide);
 showOrder=order.filter(slide=>slide>=0&&slide<state.slides.length);
 if(previousSlide!==undefined){
  if(showOrder[showCursor]===previousSlide)return;
  const candidates=showOrder.map((slide,position)=>slide===previousSlide?position:-1).filter(position=>position>=0);
  if(candidates.length){showCursor=candidates.reduce((best,position)=>Math.abs(position-showCursor)<Math.abs(best-showCursor)?position:best,candidates[0]);return;}
 }
 const current=showOrder.indexOf(index);
 if(current>=0)showCursor=current;
 else showCursor=Math.min(showCursor,Math.max(0,showOrder.length-1));
}
function showSlideAt(position,step=1,skipHidden=presenting){
 if(!showOrder.length)return -1;
 let cursor=position;
 for(let n=0;n<showOrder.length;n++){
  if(cursor<0||cursor>=showOrder.length){
   if(loopShow())cursor=(cursor+showOrder.length)%showOrder.length;
   else return -1;
  }
  const candidate=showOrder[cursor];
  if(!skipHidden||!state.hiddenSlides?.[candidate])return candidate;
  cursor+=step;
 }
 return -1;
}
function firstShowSlide(){return showSlideAt(0,1,true);}
function nextShowPosition(step){
 if(!showOrder.length)return -1;
 let cursor=showCursor+step;
 for(let n=0;n<showOrder.length;n++){
  if(cursor<0||cursor>=showOrder.length){
   if(loopShow())cursor=(cursor+showOrder.length)%showOrder.length;
   else return -1;
  }
  if(!presenting||!state.hiddenSlides?.[showOrder[cursor]])return cursor;
  cursor+=step;
 }
 return -1;
}
function nextShowSlide(step){
 const position=nextShowPosition(step);
 return position<0?-1:showOrder[position];
}
function browsePositions(){
 return showOrder.map((slide,position)=>({slide,position})).filter(({slide})=>!state.hiddenSlides?.[slide]);
}
function syncBrowseScrollbar(){
 const scrollbar=byId('presentation-scrollbar');
 const enabled=presenting&&state.showProperties?.mode?.kind==='browse'&&state.showProperties?.mode?.showScrollbar===true;
 scrollbar.setAttribute('aria-label',pt('Slide position'));
 scrollbar.hidden=!enabled;
 document.body.classList.toggle('browse-scrollbar',enabled);
 if(!enabled)return;
 const positions=browsePositions();
 const max=Math.max(0,positions.length-1);
 scrollbar.setAttribute('aria-valuemax',String(max));
 const current=positions.findIndex(({position})=>position===showCursor);
 const value=Math.max(0,current);
 scrollbar.setAttribute('aria-valuenow',String(value));
 const thumbPercent=positions.length?Math.max(8,100/positions.length):100;
 scrollbar.firstElementChild.style.height=thumbPercent+'%';
 scrollbar.firstElementChild.style.top=(max?(value/max)*(100-thumbPercent):0)+'%';
 scrollbar.setAttribute('aria-valuetext',positions.length?slideLabel(positions[value].slide):pt('No slides'));
}
function findSlide(start,step,skipHidden=presenting){
 if(skipHidden&&presenting)return showSlideAt(showCursor+step,step);
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
 const previousShowCursor=showCursor;
 const nextPosition=skipHidden&&presenting?nextShowPosition(step):-1;
 const next=skipHidden&&presenting?(nextPosition<0?-1:showOrder[nextPosition]):findSlide(index+step,step,skipHidden);
 // Arriving backwards lands on a slide that has already played out, the way
 // PowerPoint shows it; arriving forwards starts its build from the top.
 if(next>=0){
  selectSlide(next,focusThumbnail,true,step<0?'end':'start',nextPosition);
  if(presenting&&step>0&&nextPosition>=0&&nextPosition<=previousShowCursor&&loopShow())scheduleKioskRestart();
 }
 else if(presenting&&step>0&&linkedShowId!==null)void exitPresentation();
}
let advanceTimer,advanceKey;
let kioskRestartTimer,kioskRestartKey;
function clearKioskRestart(){
  clearTimeout(kioskRestartTimer);kioskRestartTimer=undefined;kioskRestartKey=null;
}
function scheduleKioskRestart(){
  clearKioskRestart();
  const restart=state.showProperties?.mode?.kind==='kiosk'?Number(state.showProperties.mode.restart):NaN;
  // The OOXML attribute is an unsigned duration, but the schema does not
  // define a special meaning for zero. Do not turn an unconfigured/zero value
  // into a busy immediate restart loop; only a positive duration schedules it.
  if(!presenting||!Number.isFinite(restart)||restart<=0)return;
  const key={};kioskRestartKey=key;
  const deadline=performance.now()+restart;
  function tick(){
    if(kioskRestartKey!==key||!presenting)return;
    const remaining=deadline-performance.now();
    if(remaining>0){kioskRestartTimer=setTimeout(tick,Math.min(remaining,2147483647));return;}
    const first=firstShowSlide();
    if(first>=0){
      const position=showOrder.indexOf(first);
      clearTimeout(advanceTimer);advanceTimer=undefined;advanceKey=null;
      selectSlide(first,false,true,'start',position);
      scheduleKioskRestart();
    }else clearKioskRestart();
  }
  kioskRestartTimer=setTimeout(tick,Math.min(restart,2147483647));
}
function scheduleAdvance(){
  const delay=state.transitions?.[index]?.advanceAfterMs;
  // A slide that still has effects to play — or one whose last click is still
  // playing out — is not finished, whatever its transition says about
  // advancing itself.
  const settled=!animationsWaiting()&&(!animationPlayer||!(animationPlayer.pending||animationPlayer.running));
  const key=presenting&&state.showProperties?.useTimings!==false&&settled&&(nextShowSlide(1)>=0||linkedShowId!==null)&&Number.isFinite(delay)&&delay>=0?index+':'+showCursor+':'+delay:null;
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
function animationsWaiting(){return state.showProperties?.showAnimation!==false&&animationLoad==='loading'&&presenting&&animationStepsAt(index).length>0;}
function animationsPending(){return state.showProperties?.showAnimation!==false&&(animationsWaiting()||Boolean(animationPlayer&&animationPlayer.pending));}
function animationsBehind(){return Boolean(animationPlayer&&animationPlayer.cursor>0);}
function updateNavigationButtons(){
  for(const id of ['prev','present-prev'])byId(id).disabled=presenting?nextShowSlide(-1)<0&&!animationsBehind():findSlide(index-1,-1)<0&&!animationsBehind();
  for(const id of ['next','present-next'])byId(id).disabled=presenting?nextShowSlide(1)<0&&linkedShowId===null&&!animationsPending():findSlide(index+1,1)<0&&!animationsPending();
  byId('present').disabled=firstShowSlide()<0;
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
let mediaPlayer,mediaPlayerKey,mediaPlayerConfigKey,createMediaPlayer,mediaLoading=false,mediaGeneration=0;
const backgroundMedia=[];
const backgroundMediaRoot=document.createElement('div');
backgroundMediaRoot.hidden=true;
stage.append(backgroundMediaRoot);
const mediaRetry=document.createElement('button');
mediaRetry.hidden=true;
byId('presentation-controls').append(mediaRetry);
mediaRetry.onclick=()=>syncMediaPlayer();
function disposeCurrentMediaPlayer(){
  mediaPlayer?.dispose();mediaPlayer=undefined;mediaPlayerKey=undefined;mediaPlayerConfigKey=undefined;
}
function disposeMediaPlayer(){
  disposeCurrentMediaPlayer();
  for(const retained of backgroundMedia)retained.player.dispose();
  backgroundMedia.length=0;
}
function leaveMediaSlide(previousIndex,previousCursor){
  if(!presenting){disposeMediaPlayer();return;}
  if(mediaPlayer){backgroundMedia.push({player:mediaPlayer,index:previousIndex,cursor:previousCursor});mediaPlayer=undefined;mediaPlayerKey=undefined;mediaPlayerConfigKey=undefined;}
  for(let i=backgroundMedia.length-1;i>=0;i--){
    const retained=backgroundMedia[i];
    if(index===retained.index||!retained.player.retainAcrossSlides(showCursor-retained.cursor,backgroundMediaRoot)){
      retained.player.dispose();backgroundMedia.splice(i,1);
    }
  }
}
function syncMediaPlayer(){
  const clips=(state.media??[]).filter(clip=>clip.slideIndex===index&&clip.kind!=='online');
  if(!presenting){disposeMediaPlayer();mediaRetry.hidden=true;return;}
  if(!clips.length){disposeCurrentMediaPlayer();mediaRetry.hidden=true;return;}
  if(!createMediaPlayer){
    if(mediaLoading)return;
    mediaLoading=true;mediaRetry.hidden=true;
    import('/media-player.js?attempt='+Date.now()).then(module=>{
      createMediaPlayer=module.createMediaPlayer;mediaLoading=false;syncMediaPlayer();
    },error=>{
      mediaLoading=false;console.warn('Could not load the media player',error);
      mediaRetry.textContent=previewLocale==='ja'?'メディアを読み込めませんでした · 再試行':'Media could not load · Retry';
      mediaRetry.hidden=!presenting;
    });
    return;
  }
  const key=JSON.stringify([index,clips]);
  if(mediaPlayerConfigKey===key)return;
  disposeCurrentMediaPlayer();
  mediaPlayer=createMediaPlayer({root:slideRoot(),overlayRoot:stage,clips,locale:previewLocale});
  mediaPlayerConfigKey=key;
  mediaPlayerKey=String(++mediaGeneration)+':'+key;
  updatePresenter();
}
function updateAnimationNotice(){
  const note=byId('present-note'),retry=byId('animation-retry');
  const slideHasEffects=presenting&&state.showProperties?.showAnimation!==false&&animationStepsAt(index).length>0;
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
 const wanted=presenting&&state.showProperties?.showAnimation!==false&&animationStepsAt(index).length>0;
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
function selectSlide(next,focusThumbnail=false,reveal=true,position='start',showPosition=-1){
 const previousIndex=index,previousCursor=showCursor;
 index=Math.max(0,Math.min(next,state.slides.length-1));
 if(showPosition>=0)showCursor=showPosition;
 else if(presenting){if(showOrder[showCursor]!==index){const naturalPosition=showOrder.indexOf(index);if(naturalPosition>=0)showCursor=naturalPosition;}}
 else {const naturalPosition=showOrder.indexOf(index);if(naturalPosition>=0)showCursor=naturalPosition;}
  if(presenting&&state.hiddenSlides?.[index]){
    const forwardPosition=nextShowPosition(1),backwardPosition=forwardPosition<0?nextShowPosition(-1):-1;
    const position=forwardPosition>=0?forwardPosition:backwardPosition;
    if(position>=0){showCursor=position;index=showOrder[position];}
  }
  if(presenting&&(previousIndex!==index||previousCursor!==showCursor))lastViewed={index:previousIndex,cursor:previousCursor};
  const count=slideCount();
  byId('chat-context').textContent=state.slides.length?count:pt('Whole project');
  byId('chat-context').dataset.focus=JSON.stringify({slide:state.slides.length?index:null,revision:state.revision??0});
  window.dispatchEvent(new Event('agent-focus'));
  if(document.body.classList.contains('editing')&&editorFocus)applyEditorFocus();
  byId('count').textContent=count;byId('present-count').textContent=count;
  document.body.classList.toggle('kiosk-presenting',presenting&&state.showProperties?.mode?.kind==='kiosk');
  byId('zoom').disabled=!state.slides.length;
  slide.hidden=!state.slides.length;byId('empty').hidden=!!state.slides.length;
  const svg=state.slides[index];
  if(previousIndex!==index||previousCursor!==showCursor)leaveMediaSlide(previousIndex,previousCursor);
  else if(svg!==displayedSvg||position!=='keep')disposeMediaPlayer();
  if(svg!==displayedSvg||previousIndex!==index){
    renderSlide(svg,presenting&&previousIndex!==index?state.transitions?.[index]:null);
  }
  syncAnimationPlayer(position);
  syncMediaPlayer();
  updateNavigationButtons();
  slide.setAttribute('aria-label',slideLabel(index));
  for(const [position,item] of Array.from(thumbnails.children).entries()){
    const button=item.firstElementChild,selected=position===index;
    button.setAttribute('aria-current',String(selected));button.tabIndex=selected?0:-1;
    button.dataset.skipped=String(!!state.hiddenSlides?.[position]);
    button.title=state.hiddenSlides?.[position]?pt('Skipped during presentation'):'';
    if(selected){if(reveal)button.scrollIntoView({block:'nearest'});if(focusThumbnail)button.focus({preventScroll:true});}
  }
  syncBrowseScrollbar();
  resize();
  if(!transitionCleanup)scheduleAdvance();
  updatePresenter();
}
function update(updated){
  const focusedThumbnail=thumbnails.contains(document.activeElement);
  const previous=state;
  if(JSON.stringify(previous.media)!==JSON.stringify(updated.media))disposeMediaPlayer();
  state=updated;
  rebuildShowOrder();
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
  if(presenting&&firstShowSlide()<0)void exitPresentation();
  selectSlide(index,focusedThumbnail,false,'keep');
  if(presenting&&(previous.showProperties?.mode?.kind!==state.showProperties?.mode?.kind||previous.showProperties?.mode?.restart!==state.showProperties?.mode?.restart))scheduleKioskRestart();
}
function setPresenting(value,from){
  if(!value){presentationFullscreen=false;finishRehearsal();}
  fullscreenRequest=null;
  clearKioskRestart();
  linkedShowId=null;showReturns=[];
  lastViewed=null;
  cancelTransition();
  rebuildShowOrder();
  presenting=value&&firstShowSlide()>=0;document.body.classList.toggle('presenting',presenting);
  if(presenting){const first=from??firstShowSlide();if(first>=0){index=first;showCursor=showOrder.findIndex(slide=>slide===first);}}
  selectSlide(index);
  if(presenting)scheduleKioskRestart();
  if(presenting)stage.focus();else{
    byId('present').focus();
    thumbnails.children[index]?.firstElementChild.scrollIntoView({block:'nearest'});
  }
}
function returnToLastViewed(){
 if(!presenting||!lastViewed||lastViewed.index>=state.slides.length)return;
 const target=lastViewed;
 selectSlide(target.index,false,true,'start',showOrder[target.cursor]===target.index?target.cursor:-1);
}
function followNavigationLink(href){
 const directions={'#pptx-next-slide':1,'#pptx-prev-slide':-1,'#pptx-first-slide':0,'#pptx-last-slide':0};
 if(!Object.hasOwn(directions,href))return false;
 const step=directions[href];
 let position=-1,target=-1;
 if(step){
  position=presenting?nextShowPosition(step):-1;
  target=presenting?(position>=0?showOrder[position]:-1):findSlide(index+step,step,true);
 }else{
  const first=href==='#pptx-first-slide';
  target=presenting?showSlideAt(first?0:showOrder.length-1,first?1:-1):findSlide(first?0:state.slides.length-1,first?1:-1,true);
  if(presenting)position=first?showOrder.indexOf(target):showOrder.lastIndexOf(target);
 }
 if(target>=0)selectSlide(target,false,true,'start',position);
 else if(presenting&&step>0)void exitPresentation();
 return true;
}
function launchCustomShow(id,returnToShow){
 if(!presenting)return;
 const show=state.customShows?.find(show=>show.id===id);
 const first=show?.slideIndices.find(slide=>slide>=0&&slide<state.slides.length&&!state.hiddenSlides?.[slide]);
 if(first===undefined)return;
 const caller={id:linkedShowId,index,cursor:showCursor,lastViewed,animation:animationPlayer?.progress};
 if(returnToShow)showReturns.push(caller);else showReturns=[];
 linkedShowId=id;showCursor=0;lastViewed=null;
 rebuildShowOrder();
 selectSlide(first,false,true,'start',showOrder.indexOf(first));
 lastViewed=null;
}
function followCustomShowLink(href){
 if(!href.startsWith('#pptx-custom-show?'))return false;
 const params=new URLSearchParams(href.slice('#pptx-custom-show?'.length));
 const id=params.get('id');
 if(id!==null&&/^[0-9]+$/.test(id))launchCustomShow(Number(id),params.get('return')==='true');
 return true;
}
// Slide Show ▸ Rehearse Timings: time each slide while presenting, then hand
// the times to the editor, which asks whether to keep them as slide timings.
let rehearsal=null;
const clock=ms=>{const total=Math.floor(ms/1000),two=n=>String(n).padStart(2,'0');return Math.floor(total/3600)+':'+two(Math.floor(total/60)%60)+':'+two(total%60);};
function rehearsalTick(){
 if(!rehearsal)return;
 const now=performance.now();
 // The show's cursor names the slide on screen; a custom show can repeat one.
 const showing=showOrder[showCursor]??index;
 if(showing!==rehearsal.slide){rehearsal.times.set(rehearsal.slide,(rehearsal.times.get(rehearsal.slide)??0)+now-rehearsal.since);rehearsal.slide=showing;rehearsal.since=now;}
 byId('rehearsal-timer').textContent=clock(now-rehearsal.since)+' · '+clock(now-rehearsal.started);
}
function startRehearsal(){
 const now=performance.now();
 rehearsal={times:new Map(),slide:showOrder[showCursor]??index,since:now,started:now,timer:setInterval(rehearsalTick,200)};
 byId('rehearsal-timer').hidden=false;rehearsalTick();
}
function finishRehearsal(){
 if(!rehearsal)return;
 rehearsalTick();
 const now=performance.now();
 rehearsal.times.set(rehearsal.slide,(rehearsal.times.get(rehearsal.slide)??0)+now-rehearsal.since);
 clearInterval(rehearsal.timer);byId('rehearsal-timer').hidden=true;
 const timings=[...rehearsal.times].map(([slide,ms])=>({slide,ms:Math.round(ms)}));
 rehearsal=null;
 editorFrame.contentWindow?.postMessage({type:'rehearsal-timings',timings},location.origin);
}
async function exitPresentation(){
  if(presenting&&showReturns.length){
    const caller=showReturns.pop();
    linkedShowId=caller.id;showCursor=caller.cursor;
    rebuildShowOrder();
    selectSlide(caller.index,false,true,'start',caller.cursor);
    lastViewed=caller.lastViewed;
    if(caller.animation&&animationPlayer)animationPlayer.resume(caller.animation.cursor,caller.animation.elapsed);
    return;
  }
  setPresenting(false);
  if(document.fullscreenElement)await document.exitFullscreen();
}
async function startPresentation(from){
  setPresenting(true,from);
  if(state.showProperties?.mode?.kind==='browse')return;
  const request=fullscreenRequest={};
  try{await document.documentElement.requestFullscreen();}
  catch{if(fullscreenRequest===request){fullscreenRequest=null;byId('exit-present').textContent=pt('Exit view · Esc');}}
}
byId('present').onclick=()=>startPresentation();
// Messages that used to sit in the shell's header now float over whichever
// view is showing, so they are not lost while the editor fills the window.
const NOTICE_MS=6000;
let noticeTimer;
function showNotice(text){const notice=byId('notice');notice.textContent=text;notice.hidden=false;clearTimeout(noticeTimer);noticeTimer=setTimeout(()=>{notice.hidden=true;},NOTICE_MS);}
function updatePresenter(){
 if(!presenterWindow||presenterWindow.closed)return;
 // The presenter runs the same player over its own copy of the slide, from the
 // same cursor and the same moment within the stop in hand, so what it shows
 // is what the audience sees — including an effect still fading in.
 const progress=animationPlayer?animationPlayer.progress:null;
 const animationSteps=animationPlayer?animationStepsAt(index):null;
 presenterWindow.postMessage({type:'presenter-state',index,count:state.slides.length,animationSteps,current:state.slides[index]??null,next:state.slides[nextShowSlide(1)]??null,hasPrevious:nextShowSlide(-1)>=0||animationsBehind(),hasNext:nextShowSlide(1)>=0||linkedShowId!==null||animationsPending(),animation:progress,notes:state.notes?.[index]??'',aspectRatio:state.aspectRatio,locale:previewLocale,presenting,media:presenting?(state.media??[]).filter(clip=>clip.slideIndex===index):[],mediaKey:mediaPlayerKey,mediaProgress:mediaPlayer?.progress??[]},location.origin);
}
setInterval(()=>{
 if(!presenterWindow||presenterWindow.closed||!mediaPlayer)return;
 presenterWindow.postMessage({type:'presenter-media-state',index,mediaKey:mediaPlayerKey,mediaProgress:mediaPlayer.progress},location.origin);
},250);
byId('presenter').onclick=()=>{
 if(presenterWindow&&!presenterWindow.closed){setPresenting(true);presenterWindow.focus();return;}
 presenterWindow=window.open('/presenter','office-kit-presenter','popup,width=1100,height=800');
 if(presenterWindow)setPresenting(true);
 else showNotice(pt('Allow popups to open presenter view.'));
};
window.addEventListener('message',event=>{
 if(event.origin!==location.origin||event.source!==presenterWindow||event.data?.type!=='presenter-command')return;
 if(event.data.action==='ready')updatePresenter();
 else if(event.data.action==='media'){
  const command=event.data.index;
  if(!presenting||!mediaPlayer||event.data.mediaKey!==mediaPlayerKey||!command||!Number.isInteger(command.shapeId)||!['play','pause','seek'].includes(command.action))return;
  if(command.action==='seek'&&(!Number.isFinite(command.time)||command.time<0))return;
  mediaPlayer.command(command);updatePresenter();
 }
 else if(event.data.action==='next')moveSlide(1,true);
 else if(event.data.action==='previous')moveSlide(-1,true);
 else if(event.data.action==='exit')void exitPresentation();
 else if(event.data.action==='lastViewed')returnToLastViewed();
 else if(event.data.action==='navigation'&&typeof event.data.index==='string')followNavigationLink(event.data.index);
 else if(event.data.action==='customShow'&&typeof event.data.index==='string')followCustomShowLink(event.data.index);
 else if(event.data.action==='jump'&&Number.isInteger(event.data.index)&&event.data.index>=0&&event.data.index<state.slides.length)selectSlide(event.data.index);
});
byId('exit-present').onclick=exitPresentation;
function scrollToBrowsePosition(value){
 const positions=browsePositions();
 const position=positions[Math.max(0,Math.min(positions.length-1,value))];
 if(position&&position.position!==showCursor)selectSlide(position.slide,false,true,'start',position.position);
}
const browseScrollbar=byId('presentation-scrollbar');
let scrollbarDragOffset;
function dragBrowseScrollbar(event){
 const rect=browseScrollbar.getBoundingClientRect();
 const thumb=browseScrollbar.firstElementChild.getBoundingClientRect();
 const travel=rect.height-thumb.height;
 if(travel<=0)return;
 const max=Number(browseScrollbar.getAttribute('aria-valuemax'));
 scrollToBrowsePosition(Math.round(((event.clientY-rect.top-scrollbarDragOffset)/travel)*max));
}
browseScrollbar.onpointerdown=event=>{
 if(event.button!==0)return;
 event.preventDefault();browseScrollbar.focus();
 const thumb=browseScrollbar.firstElementChild.getBoundingClientRect();
 scrollbarDragOffset=event.clientY>=thumb.top&&event.clientY<=thumb.bottom?event.clientY-thumb.top:thumb.height/2;
 browseScrollbar.setPointerCapture(event.pointerId);
 dragBrowseScrollbar(event);
};
browseScrollbar.onpointermove=event=>{
 if(browseScrollbar.hasPointerCapture(event.pointerId))dragBrowseScrollbar(event);
};
browseScrollbar.onpointerup=event=>{
 if(browseScrollbar.hasPointerCapture(event.pointerId))browseScrollbar.releasePointerCapture(event.pointerId);
};
browseScrollbar.onkeydown=event=>{
 const current=Number(browseScrollbar.getAttribute('aria-valuenow'));
 const max=Number(browseScrollbar.getAttribute('aria-valuemax'));
 let next;
 if(event.key==='ArrowDown'||event.key==='PageDown')next=current+1;
 else if(event.key==='ArrowUp'||event.key==='PageUp')next=current-1;
 else if(event.key==='Home')next=0;
 else if(event.key==='End')next=max;
 else return;
 event.preventDefault();event.stopPropagation();scrollToBrowsePosition(next);
};
byId('animation-retry').onclick=()=>{loadAnimationPlayer();};
loadAnimationPlayer();
document.addEventListener('fullscreenchange',()=>{
 // Only the page itself is ever requested here. A native video control can put
 // its <video> into fullscreen (and return to the page) on its own; that is the
 // viewer's choice, not a stale grant.
 if(document.fullscreenElement){
  if(document.fullscreenElement!==document.documentElement||presentationFullscreen)return;
  if(fullscreenRequest){fullscreenRequest=null;presentationFullscreen=true;}
  else void document.exitFullscreen();
 }
 else if(presentationFullscreen){presentationFullscreen=false;if(presenting)setPresenting(false);}
});
for(const id of ['prev','present-prev'])byId(id).onclick=()=>{if(!presenting||state.showProperties?.mode?.kind!=='kiosk')moveSlide(-1);};
for(const id of ['next','present-next'])byId(id).onclick=()=>{if(!presenting||state.showProperties?.mode?.kind!=='kiosk')moveSlide(1);};
byId('zoom').onchange=resize;
stage.onclick=event=>{
  const link=event.composedPath().find(node=>node instanceof Element&&node.localName==='a');
  if(link){
    const href=link.getAttribute('href')??link.getAttributeNS('http://www.w3.org/1999/xlink','href')??'';
    if(followNavigationLink(href)||followCustomShowLink(href)){event.preventDefault();return;}
    if(href==='#pptx-end-show'||href==='#pptx-last-slide-viewed'){
      event.preventDefault();
      if(!presenting)return;
      if(href==='#pptx-end-show'){void exitPresentation();return;}
      returnToLastViewed();
      return;
    }
    if(href.startsWith('#slide-')){
      event.preventDefault();
      const number=Number(href.slice(7));
      if(Number.isInteger(number)&&number>=1&&number<=state.slides.length)selectSlide(number-1);
    }
    return;
  }
  if(!presenting||state.showProperties?.mode?.kind==='kiosk'||getSelection().toString())return;
  // 'advClick' says whether a click moves to the next *slide*. A build still
  // belongs to this one, so its remaining effects play either way.
  if(animationsPending()||state.transitions?.[index]?.advanceOnClick!==false)moveSlide(1);
};
document.addEventListener('keydown',event=>{
  if(event.key==='Escape'&&presenting){event.preventDefault();void exitPresentation();return;}
  // As in PowerPoint's Reading View, Esc goes back to Normal view, unless it
  // is closing the viewer's own text field or area selection first.
  if(event.key==='Escape'&&!document.body.classList.contains('editing')&&!event.target.closest('select,input,textarea,[contenteditable]')&&document.querySelector('.slide-edit-panel')?.hidden!==false){event.preventDefault();chooseView(true);return;}
  if(presenting&&state.showProperties?.mode?.kind==='kiosk')return;
  if(event.altKey||event.ctrlKey||event.metaKey||event.target.closest('select,input,textarea,[contenteditable]'))return;
  const focusThumbnail=thumbnails.contains(document.activeElement);
  let step=0,jump=-1,jumpPosition=-1;
  if(['ArrowLeft','ArrowUp','PageUp'].includes(event.key))step=-1;
  else if(['ArrowRight','ArrowDown','PageDown'].includes(event.key))step=1;
  else if(event.key==='Home'){
    jump=presenting?firstShowSlide():findSlide(0,1);
    if(presenting)jumpPosition=showOrder.findIndex(slide=>slide===jump&&!state.hiddenSlides?.[slide]);
  }
  else if(event.key==='End'){
    jump=presenting?showSlideAt(showOrder.length-1,-1):findSlide(state.slides.length-1,-1);
    if(presenting)jumpPosition=showOrder.findLastIndex(slide=>slide===jump&&!state.hiddenSlides?.[slide]);
  }
  else if(event.key===' '&&presenting&&!event.target.closest('button,a'))step=event.shiftKey?-1:1;
  else return;
  event.preventDefault();
  // Home and End are a jump, not a step: they land on a slide with its build
  // at the start rather than walking through the one in hand.
  if(step!==0)moveSlide(step,presenting,focusThumbnail);
  else if(jump>=0)selectSlide(jump,focusThumbnail,true,'start',jumpPosition);
});
new ResizeObserver(resize).observe(stage);
let refreshId=0;
async function refresh(){
  const id=++refreshId;
  try{const response=await fetch('/state'+(state.revision===undefined?'':'?since='+state.revision));if(!response.ok)throw new Error('Preview unavailable');const updated=await response.json();if(id===refreshId){if(updated.changes){updated.slides=state.slides.slice(0,updated.count);updated.slides.length=updated.count;for(const [position,svg] of Object.entries(updated.changes))updated.slides[Number(position)]=svg;}update(updated);}}
  catch{if(id===refreshId){connectionLost=true;updatePreviewStatus();}}
}
const editorFrame=byId('editor-frame');
let editorFocus;
// The preview revision the editor last saw; it reports this apart from its focus.
let editorRevision=0;
// The Agents task pane opens on demand, like PowerPoint's panes; the editor's
// Agents button (in its tab row) mirrors that state, so every change is reported back.
const agentsKey='office-kit-agents-open';
function agentsOpen(){return !document.body.classList.contains('chat-hidden');}
function postAgentsState(focusButton=false){editorFrame.contentWindow?.postMessage({type:'host-panes',agents:agentsOpen(),focus:focusButton},location.origin);}
function setAgentsOpen(open,{focusPane=false,returnFocus=false}={}){
 document.body.classList.toggle('chat-hidden',!open);
 byId('toggle-chat').setAttribute('aria-pressed',String(open));
 try{sessionStorage.setItem(agentsKey,open?'1':'0');}catch(error){console.warn('Could not save the Agents pane state',error);}
 postAgentsState(returnFocus&&document.body.classList.contains('editing'));
 resize();
 if(open&&focusPane)byId('close-chat').focus();
 else if(returnFocus){if(document.body.classList.contains('editing'))editorFrame.focus();else byId('toggle-chat').focus();}
}
byId('toggle-chat').onclick=()=>setAgentsOpen(!agentsOpen());
byId('close-chat').onclick=()=>setAgentsOpen(false,{returnFocus:true});
try{if(sessionStorage.getItem(agentsKey)==='1')setAgentsOpen(true);}catch(error){console.warn('Could not restore the Agents pane state',error);}
function setEditorMode(editing){
 document.body.classList.toggle('editing',editing);
 if(editing&&!editorFrame.getAttribute('src'))editorFrame.src='/editor';
 if(editing&&editorFocus)applyEditorFocus();else selectSlide(index);
 resize();
}
// The rendered viewer is the editor's Reading View; Normal (or Esc) returns.
function chooseView(editing){
 setEditorMode(editing);
 try{sessionStorage.setItem('office-kit-view',editing?'editor':'preview');}catch(error){console.warn('Could not save view preference',error);}
 if(editing)editorFrame.focus();else stage.focus();
}
byId('toggle-editor').onclick=()=>chooseView(true);
try{setEditorMode(sessionStorage.getItem('office-kit-view')!=='preview');}catch(error){console.warn('Could not restore view preference',error);setEditorMode(true);}
function applyEditorFocus(){
 if(!editorFocus)return;
 index=Math.max(0,editorFocus.slide);
 byId('chat-context').textContent=(editorFocus.locale==='ja'?'スライド ':'Slide ')+(index+1)+' / '+editorFocus.count+(editorFocus.dirty?' · '+(editorFocus.locale==='ja'?'未保存':'Unsaved'):'');
 byId('chat-context').dataset.focus=JSON.stringify({slide:editorFocus.count?index:null,revision:editorFocus.dirty?-1:editorRevision});
 window.dispatchEvent(new Event('agent-focus'));
}
window.addEventListener('message',event=>{
 if(event.origin!==location.origin||event.source!==editorFrame.contentWindow||event.data?.type!=='editor-command')return;
 const slide=event.data.slide;
 if(event.data.action==='agents'){setAgentsOpen(!agentsOpen(),{focusPane:true});return;}
 if(event.data.action==='reading'){
  if(Number.isInteger(slide)&&slide>=0&&slide<state.slides.length)index=slide;
  chooseView(false);
  return;
 }
 // The editor's Slide Show tab; the click's user activation reaches this page,
 // so fullscreen and the presenter popup are still allowed.
 if(byId('present').disabled)return;
 if(event.data.action==='start')void startPresentation();
 else if(event.data.action==='current'&&Number.isInteger(slide)&&slide>=0&&slide<state.slides.length)void startPresentation(slide);
 else if(event.data.action==='presenter')byId('presenter').click();
 else if(event.data.action==='rehearse'){void startPresentation();if(presenting)startRehearsal();}
});
window.addEventListener('message',event=>{
 if(event.origin!==location.origin||event.source!==editorFrame.contentWindow||event.data?.type!=='editor-revision'||!Number.isInteger(event.data.revision))return;
 editorRevision=event.data.revision;
 if(document.body.classList.contains('editing'))applyEditorFocus();
});
window.addEventListener('message',event=>{
 if(event.origin!==location.origin||event.source!==editorFrame.contentWindow||event.data?.type!=='editor-focus')return;
 editorFocus=event.data;
 if(previewLocale!==editorFocus.locale&&(editorFocus.locale==='ja'||editorFocus.locale==='en')){previewLocale=editorFocus.locale;updatePreviewLabels();}
 updatePresenter();
 if(document.body.classList.contains('editing'))applyEditorFocus();
 postAgentsState();
});
updatePreviewLabels();
const events=new EventSource('/events');events.onmessage=event=>{if(event.data==='history'){window.dispatchEvent(new Event('agent-history'));return;}if(event.data==='chat'){window.dispatchEvent(new Event('agent-chat'));return;}if(event.data==='ready'){delete state.revision;}void refresh();};events.onerror=()=>{connectionLost=true;updatePreviewStatus()};refresh();
</script><script type="module" src="/terminal.js"></script></html>`;
