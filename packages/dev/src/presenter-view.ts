// Presenter controls share the slide-show state, including custom-show ordering.
export const presenterStyles = `
#presenter-view{display:none}.presenter-view #presenter-view{display:block;position:fixed;inset:0;z-index:12;pointer-events:none;color:#eee;font:14px -apple-system,BlinkMacSystemFont,sans-serif;background:none}
#presenter-view button,#presenter-view textarea,#presenter-view [role=separator]{pointer-events:auto}
#presenter-toolbar{position:absolute;top:0;left:0;right:0;height:54px;display:flex;align-items:center;gap:20px;padding:0 20px;background:#1b1b1b;pointer-events:auto}
#presenter-toolbar button{background:transparent;color:#eee;border:0;font:inherit}#presenter-clock{position:fixed;top:76px;right:calc(100% - var(--presenter-split) + 20px)}#presenter-timebar{position:absolute;top:65px;left:24px;display:flex;align-items:center;gap:12px}#presenter-elapsed{font:26px ui-monospace,monospace}#presenter-timebar button{border:0;background:transparent;color:#ddd;font-size:20px}
body.presenter-view{--presenter-split:65%;background:#1b1b1b}.presenter-view .workspace,.presenter-view main{background:#1b1b1b}.presenter-view #stage{position:fixed;left:20px;top:105px;bottom:165px;right:calc(100% - var(--presenter-split) + 20px);padding:0;background:#000;overflow:hidden}
#presenter-sidebar{position:absolute;top:65px;left:calc(var(--presenter-split) + 20px);right:24px;bottom:165px;display:flex;flex-direction:column;gap:12px;pointer-events:auto;min-height:0}
#presenter-next{width:100%;min-height:0;max-height:42%;object-fit:contain;object-position:left center}#presenter-next[hidden]{display:none}#presenter-next-label{margin:0;font-size:15px;font-weight:400}#presenter-notes{min-height:0;flex:1}#presenter-notes #speaker-notes{font-size:var(--presenter-note-size,22px);line-height:1.45;padding:0;color:#fff}#presenter-note-controls{display:flex;gap:10px}#presenter-note-controls button{border:0;background:transparent;font-size:18px;color:white}
#presenter-divider{position:absolute;left:calc(var(--presenter-split) - 4px);width:8px;top:65px;bottom:165px;cursor:col-resize;touch-action:none}#presenter-divider:hover,#presenter-divider:focus-visible{background:#555}
#presenter-navigation{position:absolute;bottom:0;left:0;right:0;height:120px;display:flex;align-items:center;gap:14px;overflow-x:auto;background:#1b1b1b;padding:12px 20px;pointer-events:auto}#presenter-navigation button{flex:none;width:140px;height:96px;background:none;border:2px solid transparent;color:#ddd;padding:4px;display:flex;flex-direction:column;align-items:center;gap:3px}#presenter-navigation button[aria-current=true]{border-color:#df795b}#presenter-navigation img{width:100%;height:68px;object-fit:contain}#presenter-navigation .hidden-slide span{text-decoration:line-through}
#presenter-slide-controls{position:absolute;bottom:132px;left:calc(var(--presenter-split) / 2);transform:translateX(-50%);display:flex;align-items:center;gap:16px}#presenter-slide-controls button{border:1px solid #aaa;border-radius:50%;background:transparent;color:#ddd;font-size:18px;width:26px;height:26px;padding:0}.presenter-view #presentation-controls{bottom:170px;z-index:13}.presenter-view #present-count,.presenter-view #exit-present,.presenter-view #present-prev,.presenter-view #present-next{display:none}.presenter-view #show-screen{left:20px;top:105px;right:calc(100% - var(--presenter-split) + 20px);bottom:165px}.presenter-view #browse-scrollbar{display:none}
`;
export const presenterView = `
let presenterView=false,presenterStart=0,presenterPausedAt=null,presenterInterval=null,presenterSlides=null,presenterNoteSize=22,presenterDrag=null;
const presenter=document.createElement('section');presenter.id='presenter-view';presenter.setAttribute('aria-label','Presenter View');
presenter.innerHTML='<div id="presenter-toolbar"><button id="presenter-end">End Show</button><button id="presenter-hide">Use Slide Show</button><button id="presenter-swap" hidden>Swap Displays</button><span id="presenter-display-status" role="status"></span><span id="presenter-clock"></span></div><div id="presenter-timebar"><output id="presenter-elapsed" aria-label="Elapsed time">00:00:00</output><button id="presenter-pause" aria-label="Pause timer" title="Pause timer">Ⅱ</button><button id="presenter-reset" aria-label="Reset timer" title="Reset timer">↶</button></div><div id="presenter-divider" role="separator" aria-label="Slide and notes pane width" aria-orientation="vertical" tabindex="0"></div><div id="presenter-sidebar"><h2 id="presenter-next-label">Next slide</h2><img id="presenter-next" alt="Next slide"><div id="presenter-notes"></div><div id="presenter-note-controls"><button id="presenter-notes-smaller" aria-label="Decrease notes text size">A−</button><button id="presenter-notes-larger" aria-label="Increase notes text size">A+</button></div></div><div id="presenter-slide-controls"><button id="presenter-previous" aria-label="Previous slide">‹</button><span id="presenter-position"></span><button id="presenter-advance" aria-label="Next slide">›</button></div><nav id="presenter-navigation" aria-label="Presenter slides"></nav>';
document.body.append(presenter);
presenter.addEventListener('keydown',event=>{
  if(event.target.closest('button')&&['Enter',' '].includes(event.key))event.stopPropagation();
});
function updatePresenterClock(){
  const seconds=Math.max(0,Math.floor(((presenterPausedAt??performance.now())-presenterStart)/1000));
  byId('presenter-elapsed').textContent=Math.floor(seconds/60)+':'+String(seconds%60).padStart(2,'0');
  byId('presenter-clock').textContent=new Date().toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'});
}
function nextPresenterSlide(){
  if(!customShowSequence)return visibleShowSlide(index,1);
  if(customShowPosition+1<customShowSequence.length)return customShowSlide(customShowPosition+1);
  const previous=customShowReturns.at(-1);
  if(previous)return state.editor.slides.findIndex(slide=>slide.key===previous.key);
  return showLoops()?customShowSlide(0):-1;
}
function renderPresenter(){
  if(!presenting&&presenterView)setPresenterView(false);
  if(!presenterView)return;
  const next=nextPresenterSlide(),preview=byId('presenter-next');
  preview.hidden=next<0||next>=state.slides.length;
  if(!preview.hidden){preview.src=urls[next];preview.alt='Next slide: '+(next+1);}
  else preview.removeAttribute('src');
  byId('presenter-next-label').textContent=preview.hidden?'End of slide show':'Next slide';
  const notes=byId('speaker-notes');
  if(document.activeElement!==notes)notes.value=state.editor?.slides[index]?.notes??'';
  if(presenterSlides!==state.slides){
    presenterSlides=state.slides;const nav=byId('presenter-navigation');nav.replaceChildren();
    state.slides.forEach((svg,i)=>{
      const button=document.createElement('button'),img=document.createElement('img'),label=document.createElement('span');
      button.type='button';button.dataset.slide=String(i);button.setAttribute('aria-label','Slide '+(i+1)+(state.editor?.slides[i]?.hidden?' (hidden)':''));button.classList.toggle('hidden-slide',!!state.editor?.slides[i]?.hidden);
      img.src=urls[i];img.alt='';label.textContent=String(i+1);button.append(img,label);
      button.onclick=()=>{setScreen('');selectSlide(i,false,false);stage.focus();};nav.append(button);
    });
  }
  for(const button of byId('presenter-navigation').children){const selected=Number(button.dataset.slide)===index;button.setAttribute('aria-current',String(selected));}
  byId('presenter-position').textContent=(index+1)+' / '+state.slides.length;
  byId('presenter-previous').disabled=!canPreviousShow();
  updatePresenterClock();
}
function setPresenterView(value){
  presenterView=value;document.body.classList.toggle('presenter-view',value);
  const notes=byId('speaker-notes');
  if(value){byId('presenter-notes').append(notes);presenterStart=performance.now();presenterPausedAt=null;byId('presenter-pause').setAttribute('aria-label','Pause timer');byId('presenter-pause').textContent='Ⅱ';clearInterval(presenterInterval);presenterInterval=setInterval(updatePresenterClock,1000);}
  else{closeAudience();byId('notes-pane').append(notes);clearInterval(presenterInterval);presenterInterval=null;presenterSlides=null;}
  resize();if(presenting)renderShowInk();
}
function startPresenterView(){
  if(!state.slides.length||showSettings()?.mode.kind==='kiosk'||kioskShow())return;
  if(window.screen.isExtended)openAudience();
  if(!presenting)startConfiguredShow();
  if(!presenting){closeAudience();return;}
  setPresenterView(true);
  renderPresenter();stage.focus();
}
byId('presenter-view-start').onclick=startPresenterView;
byId('presenter-end').onclick=exitPresentation;
byId('presenter-swap').onclick=swapAudienceDisplays;
byId('presenter-previous').onclick=previousShow;
byId('presenter-advance').onclick=advanceShow;
byId('presenter-hide').onclick=()=>{setPresenterView(false);stage.focus();};
byId('presenter-pause').onclick=()=>{
  const now=performance.now();
  if(presenterPausedAt===null)presenterPausedAt=now;else{presenterStart+=now-presenterPausedAt;presenterPausedAt=null;}
  const paused=presenterPausedAt!==null,button=byId('presenter-pause');button.textContent=paused?'▶':'Ⅱ';button.setAttribute('aria-label',paused?'Resume timer':'Pause timer');button.title=button.getAttribute('aria-label');updatePresenterClock();
};
byId('presenter-reset').onclick=()=>{presenterStart=performance.now();if(presenterPausedAt!==null)presenterPausedAt=presenterStart;updatePresenterClock();};
for(const [id,step] of [['presenter-notes-smaller',-2],['presenter-notes-larger',2]])byId(id).onclick=()=>{presenterNoteSize=Math.max(12,Math.min(72,presenterNoteSize+step));presenter.style.setProperty('--presenter-note-size',presenterNoteSize+'px');};
const presenterDivider=byId('presenter-divider');
function resizePresenter(value){const width=Math.max(10,Math.min(85,value));document.body.style.setProperty('--presenter-split',width+'%');presenterDivider.setAttribute('aria-valuenow',String(Math.round(width)));resize();renderShowInk();}
presenterDivider.setAttribute('aria-valuemin','10');presenterDivider.setAttribute('aria-valuemax','85');presenterDivider.setAttribute('aria-valuenow','65');
presenterDivider.onpointerdown=event=>{if(event.button!==0)return;event.preventDefault();presenterDrag=event.pointerId;presenterDivider.setPointerCapture(event.pointerId);};
presenterDivider.onpointermove=event=>{if(presenterDrag===event.pointerId)resizePresenter(event.clientX/innerWidth*100);};
for(const name of ['pointerup','pointercancel','lostpointercapture'])presenterDivider.addEventListener(name,()=>{presenterDrag=null;});
presenterDivider.onkeydown=event=>{if(!['ArrowLeft','ArrowRight'].includes(event.key))return;event.preventDefault();event.stopPropagation();resizePresenter(Number(presenterDivider.getAttribute('aria-valuenow'))+(event.key==='ArrowLeft'?-2:2));};
window.addEventListener('office-update',renderPresenter);
document.addEventListener('keydown',event=>{
  if(event.defaultPrevented||!officeMenu.hidden||event.target.closest('input,textarea,select,[contenteditable],dialog'))return;
  if(event.altKey&&!event.metaKey&&!event.ctrlKey&&!event.shiftKey&&event.key==='Enter'){event.preventDefault();event.stopImmediatePropagation();startPresenterView();}
},true);
`;
