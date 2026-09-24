// Interactions shared by the ribbon, thumbnail pane and slide show.
export const officeInteractions = `
const officeMenu=byId('office-menu');
let menuOrigin,screenMode='',slideNumber='',menuPreviewCleanup;
function clearMenuPreview(){menuPreviewCleanup?.();menuPreviewCleanup=undefined;}
officeMenu.addEventListener('pointerleave',clearMenuPreview);
function closeOfficeMenu(restore=true){clearMenuPreview();if(menuOrigin?.hasAttribute('aria-expanded'))menuOrigin.setAttribute('aria-expanded','false');officeMenu.hidden=true;officeMenu.classList.remove('shape-gallery','layout-gallery','picture-gallery','correction-gallery');officeMenu.removeAttribute('aria-label');officeMenu.replaceChildren();if(officeMenu.ownerDocument!==document)document.body.append(officeMenu);if(restore)menuOrigin?.focus({preventScroll:true});}
function positionMenu(menu,x,y){const {innerWidth,innerHeight}=menu.ownerDocument.defaultView;menu.style.left=Math.max(4,Math.min(x,innerWidth-menu.offsetWidth-4))+'px';menu.style.top=Math.max(4,Math.min(y,innerHeight-menu.offsetHeight-4))+'px';}
function fillOfficeMenu(menu,items,parentButton=null){
  const document=menu.ownerDocument,{innerWidth}=document.defaultView;
  function dismissChildren(){for(const child of menu.querySelectorAll(':scope > [role=menu]'))child.remove();for(const button of menu.querySelectorAll(':scope > button[aria-expanded]'))button.setAttribute('aria-expanded','false');}
  for(const item of items){
    if(item?.heading){const heading=document.createElement('div');heading.className='gallery-heading';heading.textContent=item.label;menu.append(heading);continue;}
    if(!item){const separator=document.createElement('div');separator.setAttribute('role','separator');menu.append(separator);continue;}
    const button=document.createElement('button');button.type='button';button.setAttribute('role','menuitem');button.tabIndex=-1;button.textContent=item.label;button.disabled=!!item.disabled;
    if(item.shortcut){const shortcut=document.createElement('span');shortcut.className='menu-shortcut';shortcut.textContent=item.shortcut;shortcut.setAttribute('aria-hidden','true');button.append(shortcut);}
    if(item.checked!==undefined){button.setAttribute('role','menuitemcheckbox');button.setAttribute('aria-checked',String(item.checked));}
    if(item.icon){
      button.title=item.label;button.setAttribute('aria-label',item.label);button.textContent='';
      const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');svg.setAttribute('viewBox','-3 -3 30 30');svg.setAttribute('aria-hidden','true');
      const path=document.createElementNS(svg.namespaceURI,'path');path.setAttribute('d',item.icon);svg.append(path);button.append(svg);
    }
    if(item.picture?.contrast!==undefined){
      button.title=item.label;button.setAttribute('aria-label',item.label);button.textContent='';
      const ns='http://www.w3.org/2000/svg',svg=document.createElementNS(ns,'svg');svg.setAttribute('viewBox','0 0 60 45');svg.setAttribute('aria-hidden','true');
      const filter=document.createElementNS(ns,'filter'),id='picture-correction-'+menu.children.length;filter.id=id;
      const transfer=document.createElementNS(ns,'feComponentTransfer');transfer.setAttribute('color-interpolation-filters','sRGB');
      for(const channel of ['R','G','B']){const fn=document.createElementNS(ns,'feFunc'+channel);fn.setAttribute('type','linear');fn.setAttribute('slope',String(1+item.picture.contrast));fn.setAttribute('intercept',String((item.picture.brightness||0)-item.picture.contrast/2));transfer.append(fn);}
      filter.append(transfer);const defs=document.createElementNS(ns,'defs');defs.append(filter);svg.append(defs);
      const img=document.createElementNS(ns,'image');img.setAttribute('href',item.picture.src);img.setAttribute('width','60');img.setAttribute('height','45');img.setAttribute('filter','url(#'+id+')');img.setAttribute('opacity',String(item.picture.opacity));svg.append(img);button.append(svg);
    }else if(item.picture){const img=document.createElement('img');img.src=item.picture.src;img.alt='';img.style.opacity=String(item.picture.opacity);button.prepend(img);}
    if(item.thumbnail){
      const svg=document.createElementNS('http://www.w3.org/2000/svg','svg');svg.setAttribute('viewBox','0 0 '+item.thumbnail.width+' '+item.thumbnail.height);svg.setAttribute('aria-hidden','true');
      for(const box of item.thumbnail.boxes){const rect=document.createElementNS(svg.namespaceURI,'rect');for(const [key,value] of Object.entries({x:box.x,y:box.y,width:box.w,height:box.h}))rect.setAttribute(key,String(value));rect.setAttribute('vector-effect','non-scaling-stroke');svg.append(rect);}
      const label=document.createElement('span');label.textContent=item.label;button.replaceChildren(svg,label);button.title=item.label;
    }
    function openChild(focus){
      dismissChildren();const child=document.createElement('div');child.className='office-menu';child.classList.toggle('shape-gallery',item.children.some(entry=>entry?.heading)&&item.children.some(entry=>entry?.icon));child.setAttribute('role','menu');child.setAttribute('aria-label',item.label);menu.append(child);fillOfficeMenu(child,item.children,button);
      const rect=button.getBoundingClientRect();positionMenu(child,rect.right+child.offsetWidth>innerWidth?rect.left-child.offsetWidth:rect.right-2,rect.top-5);button.setAttribute('aria-expanded','true');if(focus)child.querySelector('button:not(:disabled)')?.focus();
    }
    if(item.children){button.setAttribute('aria-haspopup','menu');button.setAttribute('aria-expanded','false');const arrow=document.createElement('span');arrow.textContent='›';arrow.setAttribute('aria-hidden','true');button.append(arrow);}
    button.onfocus=()=>{clearMenuPreview();if(item.preview&&!button.disabled)menuPreviewCleanup=item.preview();};
    button.onpointerenter=()=>{if(button.disabled)return;if(document.activeElement===button){clearMenuPreview();if(item.preview)menuPreviewCleanup=item.preview();}else button.focus();if(item.children)openChild(false);else dismissChildren();};
    button.onclick=()=>{if(item.children){openChild(true);return;}closeOfficeMenu();item.action();};menu.append(button);
  }
  menu.onkeydown=event=>{
    event.stopPropagation();const buttons=Array.from(menu.querySelectorAll(':scope > button:not(:disabled)'));let n=buttons.indexOf(document.activeElement);
    if(event.key==='Escape'||event.key==='Tab'){event.preventDefault();if(parentButton&&event.key==='Escape'){menu.remove();parentButton.setAttribute('aria-expanded','false');parentButton.focus();}else closeOfficeMenu();return;}
    if(!parentButton&&menuOrigin?.closest('.application-menubar')&&['ArrowLeft','ArrowRight'].includes(event.key)&&!document.activeElement.hasAttribute('aria-haspopup')){event.preventDefault();const buttons=Array.from(menuOrigin.parentElement.querySelectorAll('button'));const next=buttons[(buttons.indexOf(menuOrigin)+(event.key==='ArrowLeft'?-1:1)+buttons.length)%buttons.length];closeOfficeMenu(false);next.click();return;}
    if(event.key==='ArrowLeft'&&parentButton&&!menu.classList.contains('shape-gallery')){event.preventDefault();menu.remove();parentButton.setAttribute('aria-expanded','false');parentButton.focus();return;}
    if(event.key==='ArrowRight'&&document.activeElement.hasAttribute('aria-haspopup')){event.preventDefault();document.activeElement.click();return;}
    if((menu.classList.contains('shape-gallery')||menu.classList.contains('layout-gallery')||menu.classList.contains('picture-gallery'))&&['ArrowLeft','ArrowRight','ArrowUp','ArrowDown'].includes(event.key)){
      event.preventDefault();
      if(event.key==='ArrowLeft'||event.key==='ArrowRight'){buttons[(n+(event.key==='ArrowLeft'?-1:1)+buttons.length)%buttons.length]?.focus();return;}
      const current=buttons[n]?.getBoundingClientRect();if(!current)return;
      const direction=event.key==='ArrowDown'?1:-1;
      const rows=buttons.map(button=>({button,rect:button.getBoundingClientRect()})).filter(item=>(item.rect.top-current.top)*direction>1);
      rows.sort((a,b)=>Math.abs(a.rect.top-current.top)-Math.abs(b.rect.top-current.top)||Math.abs(a.rect.left-current.left)-Math.abs(b.rect.left-current.left));
      rows[0]?.button.focus();return;
    }
    if(event.key==='ArrowDown')n++;else if(event.key==='ArrowUp')n--;else if(event.key==='Home')n=0;else if(event.key==='End')n=buttons.length-1;else return;
    event.preventDefault();dismissChildren();buttons[(n+buttons.length)%buttons.length]?.focus();
  };
}
function showOfficeMenu(items,x,y,origin=document.activeElement,host=document.body){
  clearMenuPreview();
  if(officeMenu.parentElement!==host)host.append(officeMenu);
  if(!officeMenu.hidden&&menuOrigin===origin&&origin?.closest('.application-menubar')){closeOfficeMenu();return;}
  if(menuOrigin?.hasAttribute('aria-expanded'))menuOrigin.setAttribute('aria-expanded','false');menuOrigin=origin;if(menuOrigin?.hasAttribute('aria-expanded'))menuOrigin.setAttribute('aria-expanded','true');officeMenu.replaceChildren();officeMenu.classList.toggle('correction-gallery',origin?.dataset.edit==='picture-corrections');officeMenu.classList.toggle('picture-gallery',items.some(item=>item?.picture));officeMenu.classList.toggle('layout-gallery',items.some(item=>item?.thumbnail));officeMenu.classList.toggle('shape-gallery',items.some(item=>item?.heading));if(officeMenu.classList.contains('shape-gallery'))officeMenu.setAttribute('aria-label','Shapes');else if(officeMenu.classList.contains('layout-gallery'))officeMenu.setAttribute('aria-label',origin?.dataset.edit==='slide-layout'?'Layout':'New Slide');else officeMenu.removeAttribute('aria-label');officeMenu.hidden=false;fillOfficeMenu(officeMenu,items);positionMenu(officeMenu,x,y);officeMenu.querySelector('button:not(:disabled)')?.focus();if(presenting)renderShowPointer();
}
document.addEventListener('pointerdown',event=>{if(!officeMenu.hidden&&!officeMenu.contains(event.target)&&!(menuOrigin?.closest('.application-menubar')&&menuOrigin.contains(event.target)))closeOfficeMenu(false);});
for(const button of document.querySelectorAll('.application-menubar>button')){
  button.addEventListener('focus',()=>{for(const peer of button.parentElement.children)peer.tabIndex=peer===button?0:-1;});
  button.addEventListener('pointerenter',()=>{if(!officeMenu.hidden&&menuOrigin!==button&&menuOrigin?.closest('.application-menubar')){closeOfficeMenu(false);button.click();}});
  button.addEventListener('keydown',event=>{
    if(['ArrowDown','ArrowUp'].includes(event.key)){event.preventDefault();button.click();return;}
    if(['ArrowLeft','ArrowRight','Home','End'].includes(event.key)){event.preventDefault();const buttons=Array.from(button.parentElement.children);const index=event.key==='Home'?0:event.key==='End'?buttons.length-1:(buttons.indexOf(button)+(event.key==='ArrowLeft'?-1:1)+buttons.length)%buttons.length;buttons[index].focus();}
  });
}
window.addEventListener('resize',()=>closeOfficeMenu(false));
function setScreen(mode){if(mode)cancelTransitionPlayback();screenMode=mode;const screen=byId('show-screen');screen.hidden=!mode;screen.className=mode;screen.textContent=mode==='end'?'End of slide show, click to exit.':'';scheduleSlideAdvance();renderShowPointer();renderShowInk();}
function visibleShowSlide(from,direction){let next=from+direction;while(next>=0&&next<state.slides.length&&state.editor?.slides[next]?.hidden)next+=direction;return next;}
function canPreviousShow(){return !!shapeAnimationState?.position||(customShowSequence?customShowPosition>(customShowSlide(customShowPosition)===index?0:-1):visibleShowSlide(index,-1)>=0);}
function previousShow(){setScreen('');if(previousShapeAnimation())return;if(customShowSequence){selectCustomShowPosition(customShowPosition-(customShowSlide(customShowPosition)===index?1:0));return;}const previous=visibleShowSlide(index,-1);if(previous>=0)selectSlide(previous);}
function advanceShow(){
  if(screenMode==='end'){void exitPresentation();return;}
  if(screenMode){setScreen('');return;}
  if(advanceShapeAnimation())return;
  if(customShowSequence){if(customShowPosition+1>=customShowSequence.length){if(!returnFromCustomShow()){if(showLoops()&&customShowSequence.length){selectCustomShowPosition(0);scheduleKioskRestart();}else setScreen('end');}}else selectCustomShowPosition(customShowPosition+1);return;}
  const next=visibleShowSlide(index,1);
  if(next>=state.slides.length)setScreen('end');else selectSlide(next);
}
byId('show-screen').onclick=advanceShow;
window.addEventListener('office-update',()=>{if(presenting){byId('present-prev').disabled=!canPreviousShow();byId('present-next').disabled=false;}});
byId('present-next').onclick=advanceShow;byId('present-prev').onclick=previousShow;
function showSlideTitle(position){
  const title=state.editor?.slides[position]?.shapes.find(shape=>shape.placeholder==='title'||shape.placeholder==='ctrTitle')?.text;
  return title?.replace(/\\s+/g,' ').trim()||'Slide '+(position+1);
}
function showContextMenu(event,host=document.body,origin=null){
  if(kioskShow()){event.preventDefault();return;}
  if(!state.slides.length)return;
  event.preventDefault();
  const thumb=event.target.closest('.thumbnail');if(thumb)selectSlide(Array.from(thumbnails.children).indexOf(thumb.parentElement),true);
  const items=presenting?[
    {label:'Next',action:advanceShow},
    {label:'Previous',disabled:!canPreviousShow()&&!screenMode,action:()=>{if(screenMode)setScreen('');else previousShow();}},
    {label:'Last Viewed',disabled:lastViewed===null,action:()=>{setScreen('');selectSlide(lastViewed);}},
    null,
    {label:'By Title',children:state.slides.map((_,i)=>({label:(i+1)+' '+showSlideTitle(i),action:()=>{setScreen('');selectSlide(i);}}))},
    {label:'Pointer Options',children:showPointerItems()},
    {label:'Screen',children:[{label:'Black Screen',action:()=>setScreen('black')},{label:'White Screen',action:()=>setScreen('white')}]},
    null,{label:'End Show',action:exitPresentation}
  ]:[{label:'Zoom…',action:openZoom},{label:'Slide Show',action:()=>byId('present').click()}];
  showOfficeMenu(items,event.clientX,event.clientY,origin||thumb||stage,host);
}
thumbnails.oncontextmenu=showContextMenu;stage.oncontextmenu=showContextMenu;byId('show-screen').oncontextmenu=showContextMenu;
document.addEventListener('keydown',event=>{
  if(event.defaultPrevented||!officeMenu.hidden||event.target.closest('input,textarea,select,[contenteditable],dialog'))return;
  if(event.shiftKey&&event.key==='F10'){
    const target=document.activeElement.closest('.thumbnail')||stage,rect=target.getBoundingClientRect();
    showContextMenu({preventDefault:()=>event.preventDefault(),target,clientX:rect.left+20,clientY:rect.top+20});return;
  }
  if(!presenting){
    if(!event.altKey&&!event.ctrlKey&&((event.key==='F5'&&!event.metaKey)||(event.metaKey&&event.key==='Enter'))){event.preventDefault();startConfiguredShow(event.key==='Enter'?event.shiftKey:!event.shiftKey);}
    return;
  }
  if(kioskShow())return;
  if(event.altKey||event.ctrlKey||event.metaKey)return;
  const key=event.key.toLowerCase();
  if(key==='b'||key==='.'){event.preventDefault();setScreen(screenMode==='black'?'':'black');}
  else if(key==='w'||key===','){event.preventDefault();setScreen(screenMode==='white'?'':'white');}
  else if(key==='n'||key==='enter'){event.preventDefault();if(slideNumber){selectSlide(Number(slideNumber)-1);slideNumber='';setScreen('');}else advanceShow();}
  else if(key==='h'){event.preventDefault();if(state.editor?.slides[index+1]?.hidden){setScreen('');selectSlide(index+1);}}
  else if(key==='p'||key==='backspace'||key==='delete'){event.preventDefault();previousShow();}
  else if(/^[0-9]$/.test(key)){slideNumber=(slideNumber+key).slice(-6);event.preventDefault();}
});
const filmstripResizer=byId('filmstrip-resizer');let filmstripDrag=null;
filmstripResizer.setAttribute('aria-valuemin','100');filmstripResizer.setAttribute('aria-valuemax','480');filmstripResizer.setAttribute('aria-valuenow',String(document.querySelector('.filmstrip').getBoundingClientRect().width));
function setFilmstripWidth(width){const maximum=Math.min(480,workspace.clientWidth-240-(document.body.classList.contains('chat-hidden')?0:byId('chat').clientWidth));const value=Math.max(100,Math.min(maximum,width));workspace.style.setProperty('--filmstrip-width',value+'px');filmstripResizer.setAttribute('aria-valuenow',String(Math.round(value)));}
filmstripResizer.onpointerdown=event=>{if(event.button!==0)return;event.preventDefault();filmstripDrag={id:event.pointerId,x:event.clientX,width:document.querySelector('.filmstrip').getBoundingClientRect().width};filmstripResizer.setPointerCapture(event.pointerId);document.body.classList.add('resizing-filmstrip');};
filmstripResizer.onpointermove=event=>{if(filmstripDrag?.id===event.pointerId)setFilmstripWidth(filmstripDrag.width+event.clientX-filmstripDrag.x);};
function finishFilmstripResize(){filmstripDrag=null;document.body.classList.remove('resizing-filmstrip');}
filmstripResizer.onpointerup=finishFilmstripResize;filmstripResizer.onpointercancel=finishFilmstripResize;filmstripResizer.onlostpointercapture=finishFilmstripResize;
filmstripResizer.ondblclick=()=>workspace.style.removeProperty('--filmstrip-width');
filmstripResizer.onkeydown=event=>{if(!['ArrowLeft','ArrowRight'].includes(event.key))return;event.preventDefault();event.stopPropagation();setFilmstripWidth(document.querySelector('.filmstrip').getBoundingClientRect().width+(event.key==='ArrowLeft'?-10:10));};
`;
