// Slideshow animation state is separate from saved slide artwork.
export const animationPlayback = `
let shapeAnimationState=null;
function cancelShapeAnimationPlayback(){
  const current=shapeAnimationState;shapeAnimationState=null;
  if(!current)return;
  cancelAnimationFrame(current.frame);
  for(const [element,style] of current.original){if(style===null)element.removeAttribute('style');else element.setAttribute('style',style);}
}
function renderShapeAnimationState(time=Infinity){
  const current=shapeAnimationState;if(!current)return;
  const values=new Map(current.initial);
  for(let group=0;group<current.position;group++){
    let start=0,end=0;
    for(const effect of current.groups[group]){
      if(effect.trigger==='afterEffect')start=end;
      const delay=effect.delayMs??0,duration=effect.durationMs??0;
      const begin=start+delay;end=Math.max(end,begin+duration);
      const elapsed=group===current.position-1?time:Infinity;
      if(elapsed<begin)continue;
      const progress=duration?Math.min(1,(elapsed-begin)/duration):1;
      for(const id of effect.shapeIds){
        if(effect.effect==='appear')values.set(id,1);
        else if(effect.effect==='disappear')values.set(id,0);
        else if(effect.effect==='fadeIn')values.set(id,progress);
        else if(effect.effect==='fadeOut')values.set(id,1-progress);
      }
    }
  }
  for(const [id,value] of values)for(const element of current.targets.get(id)||[]){
    element.style.opacity=String(value);element.style.visibility=value===0?'hidden':'visible';
  }
}
function initializeShapeAnimations(previewGroups){
  cancelShapeAnimationPlayback();
  if(!presenting&&!previewGroups)return;
  const groups=previewGroups??state.editor?.slides[index]?.animations??[];
  const current={groups,position:0,initial:new Map(),targets:new Map(),original:new Map(),frame:0,running:false,autoStarted:false,preview:!!previewGroups};
  for(const effect of groups.flat())for(const id of effect.shapeIds){
    if(!effect.effect)continue;
    if(!current.initial.has(id))current.initial.set(id,['appear','fadeIn'].includes(effect.effect)?0:1);
    if(!current.targets.has(id)){
      const elements=[...canvas.querySelectorAll('[data-pptx-shape-id]')].filter(element=>element.getAttribute('data-pptx-shape-id')===id);
      current.targets.set(id,elements);for(const element of elements)current.original.set(element,element.getAttribute('style'));
    }
  }
  shapeAnimationState=current;renderShapeAnimationState();
}
function advanceShapeAnimation(){
  const current=shapeAnimationState;if(!current)return false;
  if(current.running){cancelAnimationFrame(current.frame);current.running=false;renderShapeAnimationState();scheduleSlideAdvance();return true;}
  if(current.position>=current.groups.length)return false;
  cancelSlideAdvance();current.position++;
  let start=0,end=0;
  for(const effect of current.groups[current.position-1]){
    if(effect.trigger==='afterEffect')start=end;
    end=Math.max(end,start+(effect.delayMs??0)+(effect.durationMs??0));
  }
  const began=performance.now();current.running=end>0;
  const tick=()=>{
    if(shapeAnimationState!==current)return;
    const elapsed=performance.now()-began;renderShapeAnimationState(elapsed);
    if(elapsed<end)current.frame=requestAnimationFrame(tick);
    else{current.running=false;renderShapeAnimationState();
      if(current.preview)current.frame=requestAnimationFrame(()=>{
        if(shapeAnimationState!==current)return;
        if(current.position<current.groups.length)advanceShapeAnimation();else stopAnimationPreview();
      });else scheduleSlideAdvance();
    }
  };
  tick();window.dispatchEvent(new Event('office-update'));return true;
}
function stopAnimationPreview(){
  if(!shapeAnimationState?.preview)return;
  cancelShapeAnimationPlayback();window.dispatchEvent(new Event('office-update'));
}
function previewAnimations(timingId,selectedOnly=false){
  if(presenting)return;
  let groups=state.editor?.slides[index]?.animations||[];
  const ids=new Set(Array.isArray(timingId)?timingId:timingId?[timingId]:[]);
  if(selectedOnly){
    groups=groups.map(group=>group.filter(effect=>ids.has(effect.timingId))).filter(group=>group.length);
  }else if(ids.size){
    const position=groups.findIndex(group=>group.some(effect=>ids.has(effect.timingId)));
    if(position<0)return;
    const effectIndex=groups[position].findIndex(effect=>ids.has(effect.timingId));
    groups=[groups[position].slice(effectIndex),...groups.slice(position+1)];
  }
  if(!groups.length)return;
  if(typeof cancelTransitionPlayback==='function')cancelTransitionPlayback();
  initializeShapeAnimations(groups);advanceShapeAnimation();
}
function previousShapeAnimation(){
  const current=shapeAnimationState;if(!current||!current.position)return false;
  cancelAnimationFrame(current.frame);current.running=false;current.position--;renderShapeAnimationState();scheduleSlideAdvance();window.dispatchEvent(new Event('office-update'));return true;
}
`;
