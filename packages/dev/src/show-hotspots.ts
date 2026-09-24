// PowerPoint for Mac cycles slide actions with Tab and activates them with Return.
export const showHotspots = `
function showHotspots(){
  if(!presenting||screenMode)return [];
  return [...canvas.querySelectorAll('a[href],[data-click-sound],[data-click-stop-sound]')].filter(element=>{
    if(!element.getClientRects().length)return false;
    for(let node=element;node instanceof Element;node=node.parentElement){
      const style=getComputedStyle(node);
      if(style.display==='none'||style.visibility==='hidden'||style.visibility==='collapse')return false;
    }
    return true;
  });
}
document.addEventListener('keydown',event=>{
  if(!presenting||event.defaultPrevented||!officeMenu.hidden||event.altKey||event.ctrlKey||event.metaKey)return;
  if(event.composedPath().some(node=>node instanceof Element&&node.matches('input,textarea,select,[contenteditable],dialog')))return;
  if(presenterView&&!stage.contains(event.target))return;
  const hotspots=showHotspots(),active=canvas.activeElement;
  if(event.key==='Tab'){
    event.preventDefault();event.stopImmediatePropagation();
    if(!hotspots.length)return;
    const current=hotspots.indexOf(active);
    const next=current<0?(event.shiftKey?hotspots.length-1:0):(current+(event.shiftKey?-1:1)+hotspots.length)%hotspots.length;
    const target=hotspots[next];target.setAttribute('tabindex','-1');target.focus({preventScroll:true});
  }else if(event.key==='Enter'&&!event.shiftKey&&!slideNumber&&hotspots.includes(active)){
    event.preventDefault();event.stopImmediatePropagation();
    active.dispatchEvent(new MouseEvent('click',{bubbles:true,composed:true,cancelable:true,view:window}));
  }
},true);
`;
