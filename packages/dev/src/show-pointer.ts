// Slideshow pointer state stays outside the document and editing history.
export const showPointer = `
let showPointerMode='automatic',showPointerColor='red',showPointerTimer=null,showPointerPoint=null;
const laserDot=document.createElement('div');laserDot.id='show-laser-pointer';laserDot.hidden=true;laserDot.setAttribute('aria-hidden','true');document.body.append(laserDot);
function hideShowPointer(){
  stage.style.cursor='none';slide.dataset.showCursor='none';slide.style.setProperty('--show-cursor','none');
}
function renderShowPointer(){
  clearTimeout(showPointerTimer);showPointerTimer=null;laserDot.hidden=true;
  stage.style.removeProperty('cursor');delete slide.dataset.showCursor;slide.style.removeProperty('--show-cursor');
  if(!presenting)return;
  if(showPointerMode==='hidden'||showPointerMode==='laser')hideShowPointer();
  else if(showPointerMode==='pen'){stage.style.cursor='crosshair';slide.dataset.showCursor='pen';slide.style.setProperty('--show-cursor','crosshair');}
  else if(showPointerMode==='arrow'){stage.style.cursor='default';slide.dataset.showCursor='arrow';slide.style.setProperty('--show-cursor','default');}
  else showPointerTimer=setTimeout(()=>{if(presenting&&showPointerMode==='automatic')hideShowPointer();},3000);
  if(showPointerMode==='laser'&&showPointerPoint&&!screenMode&&officeMenu.hidden){
    const box=slide.getBoundingClientRect(),{x,y}=showPointerPoint;
    if(x>=box.left&&x<=box.right&&y>=box.top&&y<=box.bottom){
      laserDot.style.left=x+'px';laserDot.style.top=y+'px';laserDot.dataset.color=showPointerColor;laserDot.hidden=false;
    }
  }
}
function setShowPointer(mode){showPointerMode=mode;renderShowPointer();renderShowInk();}
function resetShowPointer(){showPointerMode='automatic';showPointerPoint=null;renderShowPointer();}
function showPointerItems(){return [
  ...[['automatic','Automatic','⌘U'],['hidden','Hidden','⌘I'],['arrow','Arrow','⌘A'],['laser','Laser Pointer','⌘L'],['pen','Pen','⌘P']].map(([mode,label,shortcut])=>({label,shortcut,checked:showPointerMode===mode,action:()=>setShowPointer(mode)})),
  {label:'Pen Color',children:[['Red','#FF0000'],['Black','#000000'],['Blue','#0000FF'],['Green','#008000'],['Yellow','#FFFF00'],['White','#FFFFFF']].map(([label,color])=>({label,checked:showInkColor===color,action:()=>{showInkColor=color;}}))},
  {label:'Laser Color',children:['Red','Green','Blue'].map(label=>({label,checked:showPointerColor===label.toLowerCase(),action:()=>{showPointerColor=label.toLowerCase();renderShowPointer();}}))}
];}
document.addEventListener('pointermove',event=>{
  if(!presenting)return;
  showPointerPoint={x:event.clientX,y:event.clientY};renderShowPointer();
});
window.addEventListener('blur',()=>{showPointerPoint=null;laserDot.hidden=true;});
document.addEventListener('pointerout',event=>{if(!event.relatedTarget){showPointerPoint=null;laserDot.hidden=true;}});
byId('present-pointer').onclick=()=>{
  const button=byId('present-pointer'),box=button.getBoundingClientRect();
  showOfficeMenu(showPointerItems(),box.left,box.top-220,button);laserDot.hidden=true;
};
document.addEventListener('keydown',event=>{
  if(!presenting||event.defaultPrevented||kioskShow()||!officeMenu.hidden||event.target.closest('input,textarea,select,[contenteditable],dialog'))return;
  const key=event.key.toLowerCase();
  if(key==='escape'&&!event.metaKey&&!event.ctrlKey&&!event.altKey&&!event.shiftKey&&['pen','laser'].includes(showPointerMode)){
    event.preventDefault();event.stopImmediatePropagation();setShowPointer('automatic');
  }else if(event.metaKey&&!event.ctrlKey&&!event.altKey&&!event.shiftKey){
    const mode={l:'laser',a:'arrow',i:'hidden',u:'automatic',p:'pen'}[key];
    if(mode){event.preventDefault();event.stopImmediatePropagation();setShowPointer(mode);}
  }else if(event.shiftKey&&!event.metaKey&&!event.ctrlKey&&!event.altKey&&key==='e'){event.preventDefault();event.stopImmediatePropagation();eraseShowInk();
  }else if(event.ctrlKey&&!event.metaKey&&!event.altKey&&!event.shiftKey&&key==='h'){
    event.preventDefault();event.stopImmediatePropagation();setShowPointer('hidden');
  }
},true);
`;
