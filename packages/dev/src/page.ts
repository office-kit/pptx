import { animationPlayback } from './animation-playback.ts';
import { presenterAudience } from './presenter-audience.ts';
import { presenterStyles, presenterView } from './presenter-view.ts';
import { showHotspots } from './show-hotspots.ts';
import { showInk } from './show-ink.ts';
import { showPointer } from './show-pointer.ts';
import { transitionPlayback } from './transition-playback.ts';
import { editingStyles } from './editor-ui.ts';
import { previewStyles } from './styles.ts';
import { officeHeader, officeFooter } from './office-shell.ts';
import { officeInteractions } from './office-interactions.ts';
import { officeStyles } from './office-styles.ts';
export const page = `<!doctype html>
<html lang="en">
<meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>Office Kit — PowerPoint preview</title>
<link rel="stylesheet" href="/terminal.css">
<style>
${previewStyles}
#agent-workspace{position:relative;flex:1;min-height:0;min-width:0;display:flex;overflow:auto}.agent-pane{position:absolute;padding:3px 5px;display:flex;flex-direction:column;flex:1;min-width:0;min-height:0;overflow:hidden}.agent-pane iframe{border:0;width:100%;flex:1;min-height:0;border-radius:0 0 9px 9px}.agent-tools{display:flex;gap:4px;align-items:center;flex-wrap:wrap;padding:6px;background:#20253a;border:1px solid #34334e;border-radius:9px 9px 0 0}.agent-tools span{flex:1;font-size:10px;letter-spacing:.04em;color:#b3abc9}.agent-tools svg{width:14px;height:14px}.agent-pane:focus-within .agent-tools{border-color:#8b6ed577;background:#29243e}.agent-tools button{font-size:11px;padding:5px;background:transparent;border:0;display:grid;place-items:center}.agent-divider{position:absolute;z-index:2;background:#101422;cursor:col-resize;touch-action:none}.agent-divider.vertical{cursor:row-resize}.agent-divider:hover,.agent-divider:focus-visible{background:#9b87ff}.splitting-agents iframe,.resizing-chat iframe{pointer-events:none}
${officeStyles}
${editingStyles}
${presenterStyles}
</style>
${officeHeader}
<div class="workspace">
<nav class="filmstrip" aria-label="Slides"><ol id="thumbnails"></ol></nav>
<main aria-label="Slide viewer"><div id="filmstrip-resizer" role="separator" tabindex="0" aria-label="Thumbnail pane width" aria-orientation="vertical"></div><pre id="error" role="alert" hidden></pre><div id="stage" tabindex="-1"><div id="empty">Waiting for slides…</div><div id="slide" hidden></div></div><div id="notes-pane" hidden><textarea id="speaker-notes" aria-label="Slide notes" placeholder="Click to add notes"></textarea></div><div id="selection-pane" hidden></div><aside id="animation-pane" aria-label="Animation Pane" hidden><header><h2>Animation Pane</h2><button id="animation-pane-close" aria-label="Close Animation Pane">×</button></header><div class="animation-preview-controls"><button data-edit="animation-preview-from">Play From</button><button data-edit="animation-preview-selected">Play Selected</button><button data-edit="animation-preview-stop" disabled>Stop</button></div><div class="animation-order"><button data-edit="animation-earlier" aria-label="Move Earlier" title="Move Earlier" disabled>↑</button><button data-edit="animation-later" aria-label="Move Later" title="Move Later" disabled>↓</button></div><div id="animation-list" role="listbox" aria-multiselectable="true" aria-label="Animation effects"></div><button id="animation-remove" data-edit="animation-remove" disabled>Remove</button></aside><div id="edit-message" role="alert"></div></main>
<aside id="chat" aria-label="Slide chat">
<div id="chat-resizer" role="separator" tabindex="0" aria-label="Chat width" aria-orientation="vertical" aria-controls="chat" title="Drag to resize · Double-click to reset"></div>
<div class="workspace-heading"><span>Agents</span><button id="close-agents" aria-label="Close agents pane">×</button></div><div id="agent-workspace"></div><div id="chat-context" hidden></div></aside>
</div>
${officeFooter}
<div id="browse-scrollbar" tabindex="0" aria-label="Browse slides" hidden><div></div></div>
<div id="presentation-controls"><button id="present-prev" aria-label="Previous slide">‹</button><button id="present-pointer" aria-label="Pointer Options" aria-haspopup="menu">✎</button><span id="present-count"></span><button id="present-next" aria-label="Next slide">›</button><button id="exit-present">End Show</button></div>
<script>
let state={slides:[],error:null,aspectRatio:16/9},index=0,urls=[],presenting=false;
let displayedSvg,presentationTrigger,lastViewed=null;
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
${transitionPlayback}
${animationPlayback}
let customShowSequence=null,customShowPosition=0,customShowDeckOrder=true;
const customShowReturns=[];
let slideAdvanceTimer=null,slideAdvanceGeneration=0;
let kioskRestartTimer=null,kioskRestartGeneration=0,kioskRootSequence=null,kioskRootDeckOrder=true;
function cancelKioskRestart(){clearTimeout(kioskRestartTimer);kioskRestartTimer=null;kioskRestartGeneration++;}
function scheduleKioskRestart(){
  cancelKioskRestart();
  if(!kioskShow()||!kioskRootSequence?.length)return;
  const delay=showSettings().mode.restart;
  if(!Number.isFinite(delay)||delay<0)return;
  const generation=kioskRestartGeneration,deadline=performance.now()+Math.max(16,delay);
  const tick=()=>{
    if(generation!==kioskRestartGeneration||!kioskShow())return;
    const remaining=deadline-performance.now();
    if(remaining>0){kioskRestartTimer=setTimeout(tick,Math.min(remaining,2147483647));return;}
    customShowReturns.length=0;customShowSequence=[...kioskRootSequence];customShowDeckOrder=kioskRootDeckOrder;customShowPosition=0;lastViewed=null;
    stopActionSounds();selectCustomShowPosition(0);scheduleKioskRestart();
  };
  tick();
}
function cancelSlideAdvance(){clearTimeout(slideAdvanceTimer);slideAdvanceTimer=null;slideAdvanceGeneration++;}
function scheduleSlideAdvance(){
  cancelSlideAdvance();
  if(!presenting||screenMode||transitionRun||shapeAnimationState?.running)return;
  if(shapeAnimationState?.position===0&&!shapeAnimationState.autoStarted&&shapeAnimationState.groups[0]?.[0]?.trigger&&shapeAnimationState.groups[0][0].trigger!=='clickEffect'){shapeAnimationState.autoStarted=true;advanceShapeAnimation();return;}
  if(showSettings()?.useTimings===false)return;
  const delay=state.editor?.slides[index]?.transition?.advanceAfterMs;
  if(!Number.isFinite(delay)||delay<0)return;
  const generation=slideAdvanceGeneration,deadline=performance.now()+Math.max(16,delay);
  const tick=()=>{if(generation!==slideAdvanceGeneration||!presenting)return;const remaining=deadline-performance.now();if(remaining>0)slideAdvanceTimer=setTimeout(tick,Math.min(remaining,2147483647));else advanceShow();};
  tick();
}
function showSettings(){return state.editor?.showProperties;}
function kioskShow(){return presenting&&showSettings()?.mode.kind==='kiosk';}
function showLoops(){return !!showSettings()?.loop||showSettings()?.mode.kind==='kiosk';}
function configuredShowSequence(){
  const settings=showSettings(),slides=state.editor?.slides??[];
  if(settings?.slides.kind==='customShow')return [...(state.editor?.customShows.find(show=>show.id===settings.slides.id)?.slides??[])];
  const range=settings?.slides.kind==='range'?settings.slides:null;
  return slides.filter((slide,i)=>(!range||(i+1>=range.start&&i+1<=range.end))&&!slide.hidden).map(slide=>slide.key);
}
function updateBrowseScrollbar(){
  const bar=byId('browse-scrollbar'),mode=showSettings()?.mode;
  const visible=presenting&&mode?.kind==='browse'&&mode.showScrollbar&&customShowSequence?.length;
  bar.hidden=!visible;document.body.classList.toggle('browse-scrollbar-visible',!!visible);
  if(!visible)return;
  bar.firstElementChild.style.height=(bar.clientHeight*customShowSequence.length)+'px';
  bar.scrollTop=bar.clientHeight*customShowPosition;
}
byId('browse-scrollbar').addEventListener('scroll',()=>{
  const bar=byId('browse-scrollbar');
  if(bar.hidden||!customShowSequence||!bar.clientHeight)return;
  const position=Math.max(0,Math.min(customShowSequence.length-1,Math.round(bar.scrollTop/bar.clientHeight)));
  if(position!==customShowPosition)selectCustomShowPosition(position);
});
new ResizeObserver(updateBrowseScrollbar).observe(byId('browse-scrollbar'));
function enterShowFullscreen(){
  if(showSettings()?.mode.kind==='browse')return;
  document.documentElement.requestFullscreen().catch(()=>{byId('exit-present').textContent='End Show';});
}
function customShowSlide(position){
  const key=customShowSequence?.[position];
  return key===undefined?-1:(state.editor?.slides??[]).findIndex(slide=>slide.key===key);
}
function selectCustomShowPosition(position){
  const destination=customShowSlide(position);
  if(destination<0)return;
  customShowPosition=position;setScreen('');selectSlide(destination);
}
function startCustomShow(id,showAndReturn=false){
  const show=state.editor?.customShows.find(show=>show.id===id);
  if(!show?.slides.length)return;
  if(presenting&&showAndReturn)customShowReturns.push({sequence:customShowSequence,deckOrder:customShowDeckOrder,position:customShowPosition,key:state.editor?.slides[index]?.key,lastViewed});
  else customShowReturns.length=0;
  customShowSequence=[...show.slides];customShowDeckOrder=false;customShowPosition=0;
  if(presenting){selectCustomShowPosition(0);if(!showAndReturn){kioskRootSequence=[...customShowSequence];kioskRootDeckOrder=customShowDeckOrder;scheduleKioskRestart();}return;}
  presentationTrigger=document.activeElement;
  selectCustomShowPosition(0);setPresenting(true);
  enterShowFullscreen();
}
function returnFromCustomShow(){
  const previous=customShowReturns.pop();
  if(!previous)return false;
  customShowSequence=previous.sequence;customShowDeckOrder=previous.deckOrder;customShowPosition=previous.position;
  const destination=(state.editor?.slides??[]).findIndex(slide=>slide.key===previous.key);
  setScreen('');selectSlide(destination>=0?destination:0);lastViewed=previous.lastViewed;
  return true;
}
function navigateAction(href){
  const target=href.slice(7);
  const custom=target.match(/^customShow-(\\d+)-(return|exit)$/);
  if(custom){startCustomShow(Number(custom[1]),custom[2]==='return');return;}
  if(target==='endShow'){void exitPresentation();return;}
  if(presenting&&customShowSequence){
    if(target==='nextSlide'){advanceShow();return;}
    if(target==='prevSlide'){previousShow();return;}
    if(target==='firstSlide'){selectCustomShowPosition(0);return;}
    if(target==='lastSlide'){selectCustomShowPosition(customShowSequence.length-1);return;}
  }
  const destinations={nextSlide:index+1,prevSlide:index-1,firstSlide:0,lastSlide:state.slides.length-1,lastSlideViewed:lastViewed};
  const destination=Object.hasOwn(destinations,target)?destinations[target]:Number(target)-1;
  if(Number.isInteger(destination)&&destination>=0&&destination<state.slides.length)selectSlide(destination,true);
}
const actionSounds=new Set();
function stopActionSounds(){
  stopTransitionSound();
  for(const audio of actionSounds){audio.pause();audio.removeAttribute('src');audio.load();}
  actionSounds.clear();
}
function playActionSound(element,trigger){
  if(element.hasAttribute('data-'+trigger+'-stop-sound'))stopActionSounds();
  const source=element.getAttribute('data-'+trigger+'-sound');
  if(!source?.startsWith('data:audio/wav;base64,'))return;
  stopTransitionSound();
  const audio=new Audio(source);
  actionSounds.add(audio);
  const release=()=>actionSounds.delete(audio);
  audio.addEventListener('ended',release,{once:true});
  audio.addEventListener('error',release,{once:true});
  void audio.play().catch(release);
}
function externalActionUrl(href){
  if(!href)return null;
  try{const url=new URL(href,location.href);return ['http:','https:','mailto:','tel:'].includes(url.protocol)?url.href:null;}
  catch{return null;}
}
canvas.addEventListener('click',event=>{
  const link=event.composedPath().find(node=>node instanceof Element&&((node.localName==='a'&&node.hasAttribute('href'))||node.hasAttribute('data-click-sound')||node.hasAttribute('data-click-stop-sound')));
  if(!link)return;
  // Link clicks must never also advance the presentation.
  event.stopPropagation();
  const href=link.getAttribute('href');
  if(!presenting){event.preventDefault();return;}
  playActionSound(link,'click');
  if(href?.startsWith('#slide-')){
    event.preventDefault();navigateAction(href);
  }else if(!externalActionUrl(href))event.preventDefault();
});
let hoverActionPoint=null;
document.addEventListener('pointermove',event=>{
  if(hoverActionPoint&&(event.clientX!==hoverActionPoint.x||event.clientY!==hoverActionPoint.y))hoverActionPoint=null;
});
canvas.addEventListener('pointerover',event=>{
  if(!presenting||showPointerMode==='pen'||event.pointerType==='touch')return;
  if(hoverActionPoint&&event.clientX===hoverActionPoint.x&&event.clientY===hoverActionPoint.y)return;
  const action=event.composedPath().find(node=>node instanceof Element&&(node.hasAttribute('data-hover-href')||node.hasAttribute('data-hover-sound')||node.hasAttribute('data-hover-stop-sound')));
  if(!action||(event.relatedTarget instanceof Node&&action.contains(event.relatedTarget)))return;
  const href=action.getAttribute('data-hover-href');
  hoverActionPoint={x:event.clientX,y:event.clientY};
  playActionSound(action,'hover');
  if(href?.startsWith('#slide-'))navigateAction(href);
  else {const url=externalActionUrl(href);if(url)window.open(url,'_blank','noopener,noreferrer');}
});

function resize(){
  if(!state.slides.length)return;
  const style=getComputedStyle(stage);
  const width=Math.max(1,stage.clientWidth-parseFloat(style.paddingLeft)-parseFloat(style.paddingRight));
  const height=Math.max(1,stage.clientHeight-parseFloat(style.paddingTop)-parseFloat(style.paddingBottom));
  const ratio=state.aspectRatio;
  const zoom=byId('zoom').value;
  const slideWidth=presenting||zoom==='fit'?Math.min(width,height*ratio):1280*Number(zoom);
  slide.style.width=slideWidth+'px';slide.style.height=slideWidth/ratio+'px';
  const percent=Math.round(slideWidth/1280*100);byId('zoom-level').textContent=percent+'%';byId('zoom-slider').value=String(percent<=100?(percent-10)/90*1000:1000+(percent-100)/300*1000);byId('zoom-slider').setAttribute('aria-valuetext',percent+' percent');
}
function selectSlide(next,focusThumbnail=false,reveal=true){
  if(!presenting)stopTransitionSound();
  const previousSvg=canvas.querySelector('svg')?.outerHTML||displayedSvg;
  const previousIndex=index;index=Math.max(0,Math.min(next,state.slides.length-1));
  if(previousIndex!==index)lastViewed=previousIndex;
  if(presenting&&customShowSequence&&customShowSlide(customShowPosition)!==index){
    const position=customShowSequence.indexOf(state.editor?.slides[index]?.key);
    if(position>=0)customShowPosition=position;
    else if(customShowDeckOrder){
      customShowPosition=-1;
      for(let i=0;i<customShowSequence.length;i++){if(customShowSlide(i)<index)customShowPosition=i;else break;}
    }
  }
  const count=state.slides.length?'Slide '+(index+1)+' of '+state.slides.length:'No slides';
  byId('chat-context').textContent=state.slides.length?count:'Whole project';
  byId('chat-context').dataset.focus=JSON.stringify({slide:state.slides.length?index:null,revision:state.revision??0});
  window.dispatchEvent(new Event('agent-focus'));
  byId('count').textContent=count;byId('present-count').textContent=count;
  updateBrowseScrollbar();
  for(const id of ['prev','present-prev'])byId(id).disabled=index===0;
  for(const id of ['next','present-next'])byId(id).disabled=index>=state.slides.length-1;
  if(presenting){byId('present-prev').disabled=!canPreviousShow();byId('present-next').disabled=false;}
  for(const id of ['present','present-first','present-current','presenter-view-start','status-present','zoom-in','zoom-out','zoom-slider','zoom-level','zoom-fit','view-fit','view-zoom'])byId(id).disabled=!state.slides.length;
  byId('zoom').disabled=!state.slides.length;
  slide.hidden=!state.slides.length;byId('empty').hidden=!!state.slides.length;
  const svg=state.slides[index];
  const svgChanged=svg!==displayedSvg;
  if(svgChanged){
    cancelTransitionPlayback();
    displayedSvg=svg;
    canvas.innerHTML=svg?'<style>svg{display:block;width:100%;height:100%}:host([data-show-cursor]) *{cursor:var(--show-cursor)!important}</style>'+svg:'';
  }
  if(reveal||previousIndex!==index||svgChanged)initializeShapeAnimations();
  slide.setAttribute('aria-label','Slide '+(index+1));
  for(const [position,item] of Array.from(thumbnails.children).entries()){
    const button=item.firstElementChild,selected=position===index;
    button.setAttribute('aria-current',String(selected));button.tabIndex=selected?0:-1;
    if(selected){if(reveal)button.scrollIntoView({block:'nearest'});if(focusThumbnail)button.focus({preventScroll:true});}
  }
  resize();
  if(presenting&&(reveal||previousIndex!==index)){
    cancelSlideAdvance();
    playTransitionSound(state.editor?.slides[index]?.transitionSound);
    playTransition(previousSvg,state.editor?.slides[index]?.transition,scheduleSlideAdvance);
  }else if(reveal||previousIndex!==index||svgChanged){cancelTransitionPlayback();scheduleSlideAdvance();}
  window.dispatchEvent(new Event('office-update'));
}
function update(updated){
  const focusedThumbnail=!!document.activeElement?.closest('.thumbnail');
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
let editorFullscreenBeforeShow=false;
function setPresenting(value){
  if(value&&!presenting)editorFullscreenBeforeShow=!!document.fullscreenElement;
  cancelTransitionPlayback();
  hoverActionPoint=null;
  if(!value){cancelKioskRestart();kioskRootSequence=null;stopActionSounds();customShowSequence=null;customShowPosition=0;customShowReturns.length=0;}
  if(value&&!presenting){lastViewed=null;kioskRootSequence=customShowSequence?[...customShowSequence]:null;kioskRootDeckOrder=customShowDeckOrder;}
  closeOfficeMenu(false);setScreen('');slideNumber='';presenting=value;document.body.classList.toggle('presenting',value);document.body.classList.toggle('kiosk-show',kioskShow());updateBrowseScrollbar();resize();
  selectSlide(index,false,false);
  initializeShapeAnimations();
  if(value){stopTransitionSound();playTransitionSound(state.editor?.slides[index]?.transitionSound);playTransition('',state.editor?.slides[index]?.transition,scheduleSlideAdvance);}
  else scheduleSlideAdvance();
  scheduleKioskRestart();
  resetShowPointer();
  window.dispatchEvent(new Event('office-update'));
  if(value)stage.focus();else{
    (presentationTrigger?.offsetParent?presentationTrigger:byId('status-present')).focus();
    thumbnails.children[index]?.firstElementChild.scrollIntoView({block:'nearest'});
  }
}
async function exitPresentation(){
  const leaveFullscreen=!editorFullscreenBeforeShow;
  setPresenting(false);
  editorFullscreenBeforeShow=false;
  try{if(leaveFullscreen&&document.fullscreenElement)await document.exitFullscreen();}
  finally{await finishShowInk();}
}
function startConfiguredShow(fromStart=false){
  customShowSequence=configuredShowSequence();customShowDeckOrder=showSettings()?.slides.kind!=='customShow';customShowReturns.length=0;
  const range=showSettings()?.slides;
  const currentHidden=!fromStart&&customShowDeckOrder&&state.editor?.slides[index]?.hidden&&(range?.kind!=='range'||(index+1>=range.start&&index+1<=range.end));
  if(!customShowSequence.length&&!currentHidden){customShowSequence=null;return;}
  customShowPosition=fromStart?0:Math.max(0,customShowSequence.indexOf(state.editor?.slides[index]?.key));
  presentationTrigger=document.activeElement;if(!currentHidden)selectCustomShowPosition(customShowPosition);setPresenting(true);
  enterShowFullscreen();
}
byId('present').onclick=()=>startConfiguredShow();
byId('exit-present').onclick=exitPresentation;
document.addEventListener('fullscreenchange',()=>{if(!document.fullscreenElement&&presenting)void exitPresentation();});
for(const id of ['prev','present-prev'])byId(id).onclick=()=>selectSlide(index-1);
for(const id of ['next','present-next'])byId(id).onclick=()=>selectSlide(index+1);
byId('zoom').onchange=resize;
stage.onclick=event=>{if(!event.defaultPrevented&&presenting&&!getSelection().toString()&&!kioskShow()&&state.editor?.slides[index]?.transition?.advanceOnClick!==false)advanceShow();};
document.addEventListener('keydown',event=>{
  if(event.defaultPrevented||!byId('office-menu').hidden)return;
  if(presenting&&(event.key==='Escape'||(!event.altKey&&!event.ctrlKey&&!event.shiftKey&&((!event.metaKey&&event.key==='-')||(event.metaKey&&event.key==='.'))))){event.preventDefault();void exitPresentation();return;}
  if(kioskShow())return;
  if(event.altKey||event.ctrlKey||event.metaKey||event.target.closest('select,input,textarea,[contenteditable],[role=tab],dialog'))return;
  let next=index;
  if(['ArrowLeft','ArrowUp','PageUp'].includes(event.key)){if(presenting&&screenMode){event.preventDefault();setScreen('');return;}if(presenting){event.preventDefault();previousShow();return;}next--;}
  else if(['ArrowRight','ArrowDown','PageDown'].includes(event.key)){if(presenting){event.preventDefault();advanceShow();return;}next++;}
  else if(event.key==='Home'){if(presenting&&customShowSequence){event.preventDefault();selectCustomShowPosition(0);return;}next=0;}
  else if(event.key==='End'){if(presenting&&customShowSequence){event.preventDefault();selectCustomShowPosition(customShowSequence.length-1);return;}next=state.slides.length-1;}
  else if(event.key===' '&&presenting&&!event.target.closest('button,a')){if(!event.shiftKey){event.preventDefault();advanceShow();return;}event.preventDefault();previousShow();return;}
  else return;
  event.preventDefault();setScreen('');selectSlide(next,thumbnails.contains(document.activeElement));
});
new ResizeObserver(resize).observe(stage);
let refreshId=0;
async function refresh(){
  const id=++refreshId;
  try{const response=await fetch('/state'+(state.revision===undefined?'':'?since='+state.revision));if(!response.ok)throw new Error('Preview unavailable');const updated=await response.json();if(id===refreshId){if(updated.changes){updated.slides=state.slides.slice(0,updated.count);updated.slides.length=updated.count;for(const [position,svg] of Object.entries(updated.changes))updated.slides[Number(position)]=svg;}update(updated);}}
  catch{if(id===refreshId)byId('status').textContent='Reconnecting…';}
}
byId('toggle-chat').onclick=()=>{const hidden=document.body.classList.toggle('chat-hidden');byId('toggle-chat').setAttribute('aria-expanded',String(!hidden));resize();};
// Ribbon uses roving focus, like the native tab strip.
const ribbonTabs=Array.from(document.querySelectorAll('[role=tab]'));
function selectRibbon(tab){for(const item of ribbonTabs){const selected=item===tab;item.setAttribute('aria-selected',String(selected));item.tabIndex=selected?0:-1;byId(item.getAttribute('aria-controls')).hidden=!selected;}document.body.classList.remove('ribbon-collapsed');byId('collapse-ribbon').setAttribute('aria-expanded','true');byId('collapse-ribbon').setAttribute('aria-label','Collapse ribbon');window.dispatchEvent(new Event('office-ribbon-change'));}
for(const tab of ribbonTabs){tab.onclick=()=>selectRibbon(tab);tab.ondblclick=()=>byId('collapse-ribbon').click();tab.onkeydown=event=>{const visibleTabs=ribbonTabs.filter(item=>!item.hidden);let n=visibleTabs.indexOf(tab);if(event.key==='ArrowRight')n++;else if(event.key==='ArrowLeft')n--;else if(event.key==='Home')n=0;else if(event.key==='End')n=visibleTabs.length-1;else return;event.preventDefault();const next=visibleTabs[(n+visibleTabs.length)%visibleTabs.length];selectRibbon(next);next.focus();};}
byId('collapse-ribbon').onclick=()=>{const collapsed=document.body.classList.toggle('ribbon-collapsed');byId('collapse-ribbon').setAttribute('aria-expanded',String(!collapsed));byId('collapse-ribbon').setAttribute('aria-label',collapsed?'Expand ribbon':'Collapse ribbon');};
for(const id of ['home-agents','status-agents','close-agents'])byId(id).onclick=()=>byId('toggle-chat').click();
byId('home-export').onclick=()=>document.querySelector('a.download').click();
for(const id of ['present-current','status-present'])byId(id).onclick=()=>byId('present').click();
byId('present-first').onclick=()=>startConfiguredShow(true);
function setZoom(value){const select=byId('zoom'),text=String(value);if(!Array.from(select.options).some(option=>option.value===text))select.add(new Option(text,text));select.value=text;resize();}
for(const id of ['view-fit','zoom-fit'])byId(id).onclick=()=>setZoom('fit');
byId('zoom-slider').oninput=event=>{const position=Number(event.target.value);setZoom(Math.round(position<=1000?10+position/1000*90:100+(position-1000)/1000*300)/100);};
for(const [id,step] of [['zoom-out',-10],['zoom-in',10]])byId(id).onclick=()=>{const current=parseInt(byId('zoom-level').textContent);setZoom(Math.min(400,Math.max(10,Math.round(current/10)*10+step))/100);};
function openZoom(){
  const percent=parseInt(byId('zoom-level').textContent),value=byId('zoom').value==='fit'?'fit':String(percent);
  byId('zoom-percent').value=percent;
  for(const radio of document.querySelectorAll('[name=zoom-preset]'))radio.checked=radio.value===value;
  byId('zoom-dialog').showModal();
}
byId('zoom-level').onclick=openZoom;byId('view-zoom').onclick=openZoom;
byId('zoom-dialog').querySelector('form').onsubmit=event=>{if(event.submitter?.value==='apply'){const selected=document.querySelector('[name=zoom-preset]:checked');setZoom(selected?.value==='fit'?'fit':Number(byId('zoom-percent').value)/100);}};
for(const radio of document.querySelectorAll('[name=zoom-preset]'))radio.onchange=()=>{if(radio.value!=='fit')byId('zoom-percent').value=radio.value;};
byId('zoom-percent').oninput=()=>{for(const radio of document.querySelectorAll('[name=zoom-preset]'))radio.checked=false;};
byId('show-thumbnails').onchange=()=>{document.body.classList.toggle('thumbnails-hidden',!byId('show-thumbnails').checked);resize();};
for(const id of ['view-normal','status-normal'])byId(id).onclick=()=>{byId('show-thumbnails').checked=true;byId('show-thumbnails').onchange();};
${officeInteractions}
${showPointer}
${showInk}
${showHotspots}
${presenterAudience}
${presenterView}
window.office={getState:()=>({...state,index,presenting,animationPreview:!!shapeAnimationState?.preview}),selectSlide,refresh,startCustomShow,previewTransition,previewAnimations,stopAnimationPreview,menu:showOfficeMenu};
const events=new EventSource('/events');events.onmessage=event=>{if(event.data==='chat'){window.dispatchEvent(new Event('agent-chat'));return;}if(event.data==='ready'){delete state.revision;}void refresh();};events.onerror=()=>{byId('status').textContent='Reconnecting…'};refresh();
</script><script type="module" src="/editor.js"></script><script type="module" src="/terminal.js"></script></html>`;
