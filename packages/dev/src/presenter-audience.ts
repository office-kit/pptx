// The audience document contains only slide artwork and temporary show overlays.
export const presenterAudience = `
let audienceWindow=null,audienceWatch=null,audienceArtwork='',audienceArtworkStructure='',audienceObservers=[];
let audienceTransition=null,audienceAnimationFrame=null,audienceScreens=null,audienceScreen=null,audienceSwapping=false;
function updateAudienceDisplays(){
  const button=byId('presenter-swap');if(!button)return;
  button.hidden=!audienceWindow||audienceWindow.closed||!audienceScreens||audienceScreens.screens.length<2;
  button.disabled=audienceSwapping;
}
function audienceDisplayStatus(message){const status=byId('presenter-display-status');if(status)status.textContent=message;}
function placeAudience(output,target){output.moveTo(target.availLeft,target.availTop);output.resizeTo(target.availWidth,target.availHeight);audienceScreen=target;}
async function swapAudienceDisplays(){
  const output=audienceWindow,details=audienceScreens;
  if(!output||output.closed||!details||audienceSwapping)return;
  const previous=details.currentScreen;
  const target=details.screens.find(screen=>screen===audienceScreen&&screen!==previous);
  if(!target){audienceDisplayStatus('The second display is no longer available.');updateAudienceDisplays();return;}
  audienceSwapping=true;updateAudienceDisplays();audienceDisplayStatus('');
  try{
    await document.documentElement.requestFullscreen({screen:target});
    if(audienceWindow!==output||output.closed||!presenting)return;
    if(details.currentScreen!==target)throw new Error('The browser did not move Presenter View to the other display.');
    placeAudience(output,previous);renderAudience();
  }catch(error){if(audienceWindow===output)audienceDisplayStatus(error.message||'The displays could not be swapped.');}
  finally{audienceSwapping=false;updateAudienceDisplays();}
}
function closeAudience(){
  clearInterval(audienceWatch);audienceWatch=null;
  for(const observer of audienceObservers)observer.disconnect();audienceObservers=[];
  const output=audienceWindow;
  if(officeMenu.ownerDocument!==document)closeOfficeMenu(false);
  if(activeInk&&activeInk.surface.ownerDocument===output?.document)endShowInkStroke();
  if(output&&audienceAnimationFrame!==null)output.cancelAnimationFrame(audienceAnimationFrame);
  audienceAnimationFrame=null;audienceTransition=null;audienceWindow=null;audienceArtwork='';audienceArtworkStructure='';
  if(audienceScreens)audienceScreens.removeEventListener('screenschange',updateAudienceDisplays);
  audienceScreens=null;audienceScreen=null;updateAudienceDisplays();audienceDisplayStatus('');
  if(output&&!output.closed)output.close();
}
function syncAudienceTransition(){
  if(!audienceWindow||audienceWindow.closed)return;
  const frame=audienceWindow.document.getElementById('audience-slide');
  const run=presenting&&!screenMode?transitionRun:null;
  if(audienceTransition?.source!==run){
    if(audienceTransition){for(const {animation} of audienceTransition.animations)animation.cancel();audienceTransition.layer.remove();}
    audienceTransition=null;
    if(run){
      const nodes=new Map();
      const copy=node=>{
        const clone=audienceWindow.document.importNode(node,false);nodes.set(node,clone);
        for(const child of node.childNodes)clone.append(copy(child));
        if(node.shadowRoot){const root=clone.attachShadow({mode:'open'});for(const child of node.shadowRoot.childNodes)root.append(copy(child));}
        return clone;
      };
      const layer=copy(run.overlay);layer.id='audience-transition';layer.style.transformOrigin='0 0';
      frame.append(layer);
      const animations=run.animations.map(source=>{
        const target=nodes.get(source.effect.target);
        const effect=new audienceWindow.KeyframeEffect(source.effect);effect.target=target;
        const animation=new audienceWindow.Animation(effect,audienceWindow.document.timeline);
        animation.pause();return {source,animation};
      });
      audienceTransition={source:run,layer,animations,width:0,height:0,effectRevision:-1};
    }
  }
  if(audienceTransition){
    const mirror=audienceTransition,width=run.overlay.clientWidth,height=run.overlay.clientHeight;
    const revised=run.effectRevision!==mirror.effectRevision;mirror.effectRevision=run.effectRevision;
    mirror.width=width;mirror.height=height;
    mirror.layer.style.width=width+'px';mirror.layer.style.height=height+'px';
    mirror.layer.style.transform='scale('+frame.clientWidth/Math.max(1,width)+','+frame.clientHeight/Math.max(1,height)+')';
    for(const {source,animation} of mirror.animations){
      if(revised){const effect=new audienceWindow.KeyframeEffect(source.effect);effect.target=animation.effect.target;animation.effect=effect;}
      animation.currentTime=source.currentTime;
    }
  }
}
function tickAudienceTransition(){
  audienceAnimationFrame=null;
  if(!audienceWindow||audienceWindow.closed)return;
  syncAudienceTransition();
  if(audienceTransition)audienceAnimationFrame=audienceWindow.requestAnimationFrame(tickAudienceTransition);
}
function renderAudience(){
  if(!audienceWindow)return;
  if(audienceWindow.closed){closeAudience();return;}
  if(!presenting)return;
  const doc=audienceWindow.document,frame=doc.getElementById('audience-slide');
  const ratio=state.aspectRatio||16/9;
  frame.style.width=Math.min(audienceWindow.innerWidth,audienceWindow.innerHeight*ratio)+'px';
  frame.style.height=Math.min(audienceWindow.innerHeight,audienceWindow.innerWidth/ratio)+'px';
  const artwork='<style>svg{display:block;width:100%;height:100%}</style>'+(canvas.querySelector('svg')?.outerHTML||'');
  if(artwork!==audienceArtwork){
    const source=canvas.querySelector('svg'),target=frame.querySelector('#audience-artwork');
    const structure=source?.cloneNode(true);
    for(const shape of structure?.querySelectorAll('[data-pptx-shape-id]')||[])shape.removeAttribute('style');
    const signature=structure?.outerHTML||'';
    if(signature!==audienceArtworkStructure||!target.firstChild){
      target.innerHTML=artwork;audienceArtworkStructure=signature;
    }else{
      // Shape animation frames change presentation styles, not link identity.
      // Keep live audience nodes so keyboard focus survives every fade frame.
      const sourceShapes=source?.querySelectorAll('[data-pptx-shape-id]')||[];
      const targetShapes=target.querySelectorAll('[data-pptx-shape-id]');
      sourceShapes.forEach((shape,index)=>{const style=shape.getAttribute('style');if(style===null)targetShapes[index].removeAttribute('style');else targetShapes[index].setAttribute('style',style);});
    }
    audienceArtwork=artwork;
  }
  syncAudienceTransition();
  if(audienceTransition&&audienceAnimationFrame===null)audienceAnimationFrame=audienceWindow.requestAnimationFrame(tickAudienceTransition);
  const ink=doc.getElementById('audience-ink');
  ink.setAttribute('viewBox',inkLayer.getAttribute('viewBox')||'0 0 1 1');
  ink.innerHTML=inkLayer.innerHTML;
  const screen=doc.getElementById('audience-screen');
  screen.hidden=!screenMode;screen.style.background=screenMode==='white'?'white':'black';
  screen.textContent=screenMode==='end'?'End of slide show, click to exit.':'';
  const cursor=slide.style.getPropertyValue('--show-cursor');
  doc.body.style.cursor=cursor||'auto';
  frame.style.setProperty('--audience-cursor',cursor||'auto');
  frame.toggleAttribute('data-force-cursor',!!cursor);
  const laser=doc.getElementById('audience-laser'),box=slide.getBoundingClientRect();
  laser.hidden=laserDot.hidden||!!screenMode;
  if(!laser.hidden&&showPointerPoint){
    laser.style.left=((showPointerPoint.x-box.left)/box.width*100)+'%';
    laser.style.top=((showPointerPoint.y-box.top)/box.height*100)+'%';
    laser.style.background={red:'#ff3030',green:'#00ff44',blue:'#3080ff'}[showPointerColor];
  }
}
const audienceHotspotSelector='a[href],[data-click-sound],[data-click-stop-sound]';
const audienceHoverSelector='[data-hover-href],[data-hover-sound],[data-hover-stop-sound]';
function audienceSourceHotspot(element,selector=audienceHotspotSelector){
  const artwork=audienceWindow?.document.getElementById('audience-artwork');
  const target=element?.closest?.(selector);
  if(!target||!artwork?.contains(target))return null;
  const position=[...artwork.querySelectorAll(selector)].indexOf(target);
  return canvas.querySelectorAll(selector)[position]??null;
}
function focusAudienceHotspot(){
  if(!audienceWindow||audienceWindow.closed)return;
  const position=[...canvas.querySelectorAll(audienceHotspotSelector)].indexOf(canvas.activeElement);
  const target=audienceWindow.document.getElementById('audience-artwork').querySelectorAll(audienceHotspotSelector)[position];
  if(target){target.setAttribute('tabindex','-1');target.focus({preventScroll:true});}
}
function openAudience(){
  if(audienceWindow&&!audienceWindow.closed)return;
  const output=window.open('','_blank','popup,width=960,height=540');
  if(!output){audienceDisplayStatus('The slide show window was blocked. Allow pop-ups, then start Presenter View again.');return;}
  audienceWindow=output;audienceArtwork='';audienceArtworkStructure='';
  const doc=output.document;doc.title='Slide Show';
  doc.head.innerHTML='<meta name="viewport" content="width=device-width, initial-scale=1"><style>html,body{margin:0;width:100%;height:100%;overflow:hidden;background:#000;color:#fff;font:18px -apple-system,sans-serif}body{display:flex;align-items:center;justify-content:center;cursor:none}#audience-slide{position:relative;overflow:hidden;z-index:0}#audience-artwork,#audience-ink{position:absolute;inset:0;width:100%;height:100%}#audience-ink{z-index:1001;pointer-events:none}#audience-artwork>svg{width:100%;height:100%;display:block}#audience-screen{position:fixed;inset:0;padding-top:30px;text-align:center;z-index:3}#audience-screen[hidden],#audience-laser[hidden]{display:none}#audience-slide[data-force-cursor] *{cursor:var(--audience-cursor)!important}#audience-laser{pointer-events:none;position:absolute;width:12px;height:12px;border-radius:50%;transform:translate(-50%,-50%);box-shadow:0 0 8px currentColor;z-index:1002}#audience-artwork a[href],#audience-artwork [data-click-sound],#audience-artwork [data-click-stop-sound]{cursor:pointer}#audience-artwork a:focus{outline:auto}</style>';
  doc.body.innerHTML='<main id="audience-slide" aria-label="Slide show"><div id="audience-artwork"></div><svg xmlns="http://www.w3.org/2000/svg" id="audience-ink"></svg><div id="audience-laser" hidden></div></main><div id="audience-screen" hidden></div>';
  const menuStyle=doc.createElement('style');
  menuStyle.textContent=[...document.styleSheets].flatMap(sheet=>{try{return [...sheet.cssRules].filter(rule=>rule.selectorText?.includes('.office-menu')).map(rule=>rule.cssText);}catch{return [];}}).join('')+'.office-menu{cursor:default}.office-menu button{color:inherit;cursor:default}';
  doc.head.append(menuStyle);
  const audienceSlide=doc.getElementById('audience-slide');audienceSlide.tabIndex=-1;
  bindShowInkSurface(audienceSlide);
  doc.addEventListener('contextmenu',event=>{
    event.preventDefault();if(!presenting)return;
    showContextMenu(event,doc.body,audienceSlide);renderAudience();
  });
  doc.addEventListener('pointerdown',event=>{if(!officeMenu.hidden&&!officeMenu.contains(event.target))closeOfficeMenu(false);});
  output.addEventListener('resize',()=>{if(officeMenu.ownerDocument===doc)closeOfficeMenu(false);renderAudience();});
  const clearAudiencePointer=()=>{showPointerPoint=null;laserDot.hidden=true;renderAudience();};
  output.addEventListener('blur',clearAudiencePointer);
  doc.addEventListener('pointerout',event=>{if(!event.relatedTarget)clearAudiencePointer();});
  doc.addEventListener('pointermove',event=>{
    if(presenting){
      const source=slide.getBoundingClientRect(),target=doc.getElementById('audience-slide').getBoundingClientRect();
      showPointerPoint={x:source.left+(event.clientX-target.left)/target.width*source.width,y:source.top+(event.clientY-target.top)/target.height*source.height};
      renderShowPointer();renderAudience();
    }
    if(hoverActionPoint&&(event.clientX!==hoverActionPoint.x||event.clientY!==hoverActionPoint.y))hoverActionPoint=null;
  });
  doc.addEventListener('pointerover',event=>{
    if(!presenting||screenMode||showPointerMode==='pen'||event.pointerType==='touch')return;
    const target=audienceSourceHotspot(event.target,audienceHoverSelector);
    if(!target)return;
    const relatedTarget=audienceSourceHotspot(event.relatedTarget,audienceHoverSelector);
    target.dispatchEvent(new PointerEvent('pointerover',{bubbles:true,composed:true,pointerType:event.pointerType,clientX:event.clientX,clientY:event.clientY,relatedTarget}));
    renderAudience();
  });
  doc.addEventListener('click',event=>{
    if(event.composedPath().includes(officeMenu))return;
    event.preventDefault();
    if(screenMode==='end'){exitPresentation();return;}
    if(!screenMode&&showPointerMode!=='pen'){
      const target=audienceSourceHotspot(event.target);
      if(target){target.dispatchEvent(new MouseEvent('click',{bubbles:true,composed:true,cancelable:true,view:window}));renderAudience();return;}
    }
    if(showPointerMode!=='pen')stage.dispatchEvent(new MouseEvent('click',{bubbles:true,cancelable:true}));
  });
  doc.addEventListener('keydown',event=>{
    if(event.shiftKey&&event.key==='F10'){
      const box=audienceSlide.getBoundingClientRect();
      showContextMenu({preventDefault:()=>event.preventDefault(),target:audienceSlide,clientX:box.left+20,clientY:box.top+20},doc.body,audienceSlide);renderAudience();return;
    }
    if(['Tab','Enter',' ','Escape','ArrowLeft','ArrowRight','ArrowUp','ArrowDown','PageUp','PageDown','Home','End','Backspace','Delete','.','-',','].includes(event.key)||event.key.length===1){
      event.preventDefault();stage.dispatchEvent(new KeyboardEvent('keydown',{key:event.key,code:event.code,metaKey:event.metaKey,ctrlKey:event.ctrlKey,altKey:event.altKey,shiftKey:event.shiftKey,bubbles:true,cancelable:true}));
      renderAudience();if(event.key==='Tab')focusAudienceHotspot();
    }
  });
  for(const target of [canvas,inkLayer,laserDot,byId('show-screen')]){
    const observer=new MutationObserver(renderAudience);observer.observe(target,{subtree:true,childList:true,attributes:true,characterData:true});audienceObservers.push(observer);
  }
  audienceWatch=setInterval(renderAudience,250);
  audienceAnimationFrame=output.requestAnimationFrame(tickAudienceTransition);
  renderAudience();
  // Placement requires a browser-granted Window Management permission.
  if(window.getScreenDetails)window.getScreenDetails().then(details=>{
    if(audienceWindow!==output||output.closed)return;
    audienceScreens=details;details.addEventListener('screenschange',updateAudienceDisplays);
    const target=details.screens.find(screen=>screen!==details.currentScreen);
    if(target)placeAudience(output,target);
    updateAudienceDisplays();
  }).catch(()=>{});
}
window.addEventListener('office-update',renderAudience);
window.addEventListener('pagehide',closeAudience);
`;
