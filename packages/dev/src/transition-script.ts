/** Presentation-only visual transitions. All layers stay inside the clipped slide. */
export const transitionScript = `
let transitionAnimations=[],transitionCleanup;
const reducedMotion=matchMedia('(prefers-reduced-motion: reduce)');
function cancelTransition(){
  const cleanup=transitionCleanup;transitionCleanup=undefined;
  for(const animation of transitionAnimations)animation.cancel();
  transitionAnimations=[];cleanup?.();
}
reducedMotion.addEventListener('change',()=>{if(reducedMotion.matches){cancelTransition();scheduleAdvance();}});
function renderSlide(svg,options){
  const previous=displayedSvg;
  cancelTransition();
  displayedSvg=svg;
  const style='<style>svg{display:block;width:100%;height:100%}.transition-layer{position:absolute;inset:0;background:white;overflow:hidden}.transition-old{pointer-events:none}</style>';
  canvas.innerHTML=svg?style+svg:'';
  const effect=options?.effect;
  if(!previous||!svg||reducedMotion.matches||(effect==='cut'&&!options.thruBlack)||!['cut','fade','push','wipe','cover','pull','zoom','split','circle','diamond','plus','blinds','comb','checker'].includes(effect))return;
  const incoming=document.createElement('div'),outgoing=document.createElement('div');
  incoming.className='transition-layer';incoming.innerHTML=svg;
  outgoing.className='transition-layer transition-old';outgoing.innerHTML=previous;
  outgoing.setAttribute('aria-hidden','true');outgoing.inert=true;
  canvas.innerHTML=style;canvas.append(outgoing,incoming);
  const duration=options.speed==='slow'?1000:options.speed==='fast'?300:600;
  const timing={duration,easing:'ease-in-out',fill:'both'};
  const animate=(node,frames,easing=timing.easing)=>transitionAnimations.push(node.animate(frames,{...timing,easing}));
  const vectors={l:[-100,0],r:[100,0],u:[0,-100],d:[0,100],lu:[-100,-100],ru:[100,-100],ld:[-100,100],rd:[100,100]};
  const [x,y]=vectors[options.direction]??vectors.l;
  const shift=(a,b)=>'translate('+a+'%,'+b+'%)';
  if(effect==='cut'){
    slide.style.background='black';outgoing.style.visibility='hidden';
    animate(incoming,[{opacity:0},{opacity:1}],'steps(1,end)');
  }else if(effect==='fade'){
    if(options.thruBlack){
      outgoing.style.background='black';incoming.style.background='black';slide.style.background='black';
      animate(outgoing,[{opacity:1,offset:0},{opacity:0,offset:.5},{opacity:0,offset:1}]);
      animate(incoming,[{opacity:0,offset:0},{opacity:0,offset:.5},{opacity:1,offset:1}]);
    }else animate(incoming,[{opacity:0},{opacity:1}]);
  }else if(effect==='push'||effect==='cover'){
    animate(incoming,[{transform:shift(-x,-y)},{transform:shift(0,0)}]);
    if(effect==='push')animate(outgoing,[{transform:shift(0,0)},{transform:shift(x,y)}]);
  }else if(effect==='pull'){
    canvas.append(outgoing);
    animate(outgoing,[{transform:shift(0,0)},{transform:shift(x,y)}]);
  }else if(effect==='wipe'){
    const clips={l:'inset(0 0 0 100%)',r:'inset(0 100% 0 0)',u:'inset(100% 0 0 0)',d:'inset(0 0 100% 0)'};
    animate(incoming,[{clipPath:clips[options.direction]??clips.l},{clipPath:'inset(0 0 0 0)'}]);
  }else if(effect==='blinds'||effect==='comb'){
    const vertical=options.direction==='vert';
    const bands=(progress)=>{
      const points=[];
      for(let band=0;band<8;band++){
        const start=band*12.5,end=start+12.5;
        const left=effect==='comb'&&band%2?100*(1-progress):0;
        const right=effect==='comb'&&band%2?100:100*progress;
        const rect=effect==='blinds'
          ?[[start,0],[start+12.5*progress,0],[start+12.5*progress,100],[start,100]]
          :[[left,start],[right,start],[right,end],[left,end]];
        // Return along the same bridge so disconnected bands share one clip polygon.
        points.push([0,0],...rect,rect[0],[0,0]);
      }
      return 'polygon('+points.map(([x,y])=>vertical?y+'% '+x+'%':x+'% '+y+'%').join(',')+')';
    };
    animate(incoming,[{clipPath:bands(0)},{clipPath:bands(1)}]);
  }else if(effect==='checker'){
    const cells=(progress)=>{
      const points=[];
      for(let row=0;row<6;row++)for(let col=0;col<8;col++){
        const amount=Math.max(0,Math.min(1,2*progress-(row+col)%2));
        const left=col*100/8,top=row*100/6;
        const right=(col+(options.direction==='vert'?1:amount))*100/8;
        const bottom=(row+(options.direction==='vert'?amount:1))*100/6;
        // Keep zero-area bridges on cell boundaries to avoid diagonal antialiasing artifacts.
        points.push([0,0],[left,0],[left,top],[right,top],[right,bottom],[left,bottom],[left,top],[left,0],[0,0]);
      }
      return 'polygon('+points.map(([x,y])=>x+'% '+y+'%').join(',')+')';
    };
    animate(incoming,[{clipPath:cells(0)},{clipPath:cells(.5)},{clipPath:cells(1)}]);
  }else if(effect==='split'){
    const collapsed=options.orientation==='vert'?'inset(0 50%)':'inset(50% 0)';
    const expanded='inset(0 0)';
    if(options.direction==='in'){
      canvas.append(outgoing);
      animate(outgoing,[{clipPath:expanded},{clipPath:collapsed}]);
    }else animate(incoming,[{clipPath:collapsed},{clipPath:expanded}]);
  }else if(effect==='circle'){
    animate(incoming,[{clipPath:'circle(0% at 50% 50%)'},{clipPath:'circle(100% at 50% 50%)'}]);
  }else if(effect==='diamond'){
    animate(incoming,[{clipPath:'polygon(50% 50%,50% 50%,50% 50%,50% 50%)'},{clipPath:'polygon(50% -50%,150% 50%,50% 150%,-50% 50%)'}]);
  }else if(effect==='plus'){
    const cross=(low,high)=>'polygon('+low+'% 0%,'+high+'% 0%,'+high+'% '+low+'%,100% '+low+'%,100% '+high+'%,'+high+'% '+high+'%,'+high+'% 100%,'+low+'% 100%,'+low+'% '+high+'%,0% '+high+'%,0% '+low+'%,'+low+'% '+low+'%)';
    animate(incoming,[{clipPath:cross(50,50)},{clipPath:cross(0,100)}]);
  }else if(effect==='zoom'){
    animate(incoming,[{transform:options.direction==='out'?'scale(2)':'scale(0.1)',opacity:0},{transform:'scale(1)',opacity:1}]);
  }
  clearTimeout(advanceTimer);advanceKey=null;
  const cleanup=()=>{outgoing.remove();incoming.remove();canvas.innerHTML=style+svg;slide.style.background='';};
  transitionCleanup=cleanup;
  Promise.all(transitionAnimations.map(animation=>animation.finished)).then(()=>{
    if(transitionCleanup!==cleanup)return;
    cancelTransition();scheduleAdvance();
  },()=>{});
}
`;
