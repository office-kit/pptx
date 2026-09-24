// Temporary slideshow ink is committed only when the presenter chooses Keep.
export const showInk = `
let showInkColor='#FF0000',activeInk=null,inkExitPending=null,suppressInkClick=null;
const showInkStrokes=new Map();
const inkLayer=document.createElementNS('http://www.w3.org/2000/svg','svg');
inkLayer.id='show-ink';inkLayer.setAttribute('aria-hidden','true');
inkLayer.style.cssText='position:fixed;pointer-events:none;z-index:9;overflow:hidden;';
document.body.append(inkLayer);
function renderShowInk(){
  if(activeInk&&(!presenting||screenMode||showPointerMode!=='pen'||activeInk.key!==state.editor?.slides[index]?.key))endShowInkStroke();
  const box=slide.getBoundingClientRect(),model=state.editor;
  inkLayer.style.display=presenting&&!screenMode?'block':'none';
  if(!model)return;
  inkLayer.style.left=box.left+'px';inkLayer.style.top=box.top+'px';inkLayer.style.width=box.width+'px';inkLayer.style.height=box.height+'px';
  inkLayer.setAttribute('viewBox','0 0 '+model.width+' '+model.height);
  inkLayer.replaceChildren();
  for(const stroke of showInkStrokes.get(model.slides[index]?.key)||[]){
    const points=stroke.points,path=document.createElementNS(inkLayer.namespaceURI,'path');
    path.setAttribute('d',points.map((p,i)=>(i?'L':'M')+p.x+' '+p.y).join(' ')+(points.length===1?' l0 0':''));
    path.setAttribute('fill','none');path.setAttribute('stroke',stroke.color);path.setAttribute('stroke-width',stroke.widthEmu);
    path.setAttribute('stroke-linecap','round');path.setAttribute('stroke-linejoin','round');inkLayer.append(path);
  }
}
function endShowInkStroke(){const current=activeInk;activeInk=null;if(current&&current.surface.hasPointerCapture(current.id))current.surface.releasePointerCapture(current.id);}
function eraseShowInk(){showInkStrokes.delete(state.editor?.slides[index]?.key);endShowInkStroke();renderShowInk();}
function inkPoint(event,surface){const box=surface.getBoundingClientRect();const x=Math.max(0,Math.min(1,(event.clientX-box.left)/box.width));const y=Math.max(0,Math.min(1,(event.clientY-box.top)/box.height));return {x:Math.round(x*state.editor.width),y:Math.round(y*state.editor.height)};}
function bindShowInkSurface(surface){
const inkDocument=surface.ownerDocument;
inkDocument.addEventListener('pointerdown',()=>{suppressInkClick=null;},true);
inkDocument.addEventListener('click',event=>{
  if(suppressInkClick===null||event.detail===0)return;
  if(event.pointerId!==undefined&&event.pointerId!==suppressInkClick)return;
  suppressInkClick=null;event.preventDefault();event.stopImmediatePropagation();
},true);
surface.addEventListener('pointerdown',event=>{
  if(!presenting||showPointerMode!=='pen'||screenMode||!officeMenu.hidden||event.button!==0||kioskShow())return;
  event.preventDefault();event.stopImmediatePropagation();suppressInkClick=event.pointerId;endShowInkStroke();
  const key=state.editor.slides[index].key,stroke={points:[inkPoint(event,surface)],color:showInkColor,widthEmu:25400};
  if(!showInkStrokes.has(key))showInkStrokes.set(key,[]);
  showInkStrokes.get(key).push(stroke);activeInk={id:event.pointerId,key,stroke,surface};surface.setPointerCapture(event.pointerId);renderShowInk();
},true);
surface.addEventListener('pointermove',event=>{
  if(!activeInk||activeInk.surface!==surface||activeInk.id!==event.pointerId)return;
  event.preventDefault();event.stopImmediatePropagation();
  if(activeInk.key!==state.editor.slides[index]?.key){endShowInkStroke();return;}
  for(const sample of event.getCoalescedEvents?.().length?event.getCoalescedEvents():[event]){
    const point=inkPoint(sample,surface),last=activeInk.stroke.points.at(-1);
    if(point.x!==last.x||point.y!==last.y)activeInk.stroke.points.push(point);
  }
  renderShowInk();
},true);
for(const name of ['pointerup','pointercancel','lostpointercapture'])surface.addEventListener(name,event=>{
  if(activeInk?.surface!==surface||activeInk.id!==event.pointerId)return;
  if(name==='pointerup'&&activeInk.key===state.editor.slides[index]?.key){
    const point=inkPoint(event,surface),last=activeInk.stroke.points.at(-1);
    if(point.x!==last.x||point.y!==last.y)activeInk.stroke.points.push(point);
  }
  endShowInkStroke();event.stopPropagation();renderShowInk();
});
surface.addEventListener('click',event=>{if(presenting&&showPointerMode==='pen'){event.preventDefault();event.stopImmediatePropagation();}},true);
}
bindShowInkSurface(slide);
window.addEventListener('office-update',renderShowInk);
new ResizeObserver(renderShowInk).observe(slide);
async function finishShowInk(){
  if(inkExitPending)return inkExitPending;
  endShowInkStroke();
  if(!showInkStrokes.size)return;
  inkExitPending=new Promise(resolve=>{
    const dialog=document.createElement('dialog');dialog.className='section-dialog';dialog.setAttribute('aria-label','Keep ink annotations');
    dialog.innerHTML='<h2>Do you want to keep your ink annotations?</h2><p role="alert"></p><footer><button data-discard>Discard</button><button data-keep autofocus>Keep</button></footer>';
    const finish=()=>{showInkStrokes.clear();renderShowInk();dialog.close();dialog.remove();resolve();};
    dialog.oncancel=event=>event.preventDefault();
    dialog.querySelector('[data-discard]').onclick=finish;
    dialog.querySelector('[data-keep]').onclick=async()=>{
      const buttons=[...dialog.querySelectorAll('button')];buttons.forEach(button=>button.disabled=true);
      try{
        const ink=[...showInkStrokes].map(([slide,strokes])=>({slide,strokes}));
        const response=await fetch('/edit',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({revision:state.revision,command:{type:'ink',slide:index,ink}})});
        const result=await response.json();if(!response.ok)throw new Error(result.error||'Could not keep ink annotations.');
        await refresh();finish();
      }catch(error){dialog.querySelector('[role="alert"]').textContent=error.message;buttons.forEach(button=>button.disabled=false);}
    };
    document.body.append(dialog);dialog.showModal();
  });
  await inkExitPending;inkExitPending=null;
}
`;
