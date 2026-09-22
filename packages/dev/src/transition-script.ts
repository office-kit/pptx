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
  if(!previous||!svg||reducedMotion.matches||(effect==='cut'&&!options.thruBlack)||!['cut','fade','push','wipe','cover','pull','zoom','split','circle','diamond','plus','blinds','comb','checker','strips','randomBar','dissolve'].includes(effect))return;
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
  }else if(effect==='dissolve'){
    const columns=32,stages=16,bounds=incoming.getBoundingClientRect();
    const rows=Math.max(1,Math.round(columns*bounds.height/bounds.width)),count=columns*rows;
    const order=Array.from({length:count},(_,i)=>i);
    for(let i=count-1;i>0;i--){
      const j=Math.floor(Math.random()*(i+1));
      [order[i],order[j]]=[order[j],order[i]];
    }
    const cells=(step)=>{
      const points=['0% 0%'];
      for(let i=0;i<Math.floor(count*step/stages);i++){
        const cell=order[i],col=cell%columns,row=Math.floor(cell/columns);
        const left=100*col/columns,right=100*(col+1)/columns,top=100*row/rows,bottom=100*(row+1)/rows;
        // Retrace bridges along cell boundaries to avoid diagonal clipping artifacts.
        for(const [x,y] of [[left,0],[left,top],[right,top],[right,bottom],[left,bottom],[left,top],[left,0],[0,0]])
          points.push(x+'% '+y+'%');
      }
      return step===0?'polygon(0% 0%,0% 0%,0% 0%)':'polygon('+points.join(',')+')';
    };
    // Discrete masks add whole squares instead of morphing unrelated polygon vertices.
    animate(incoming,Array.from({length:stages+1},(_,step)=>({clipPath:cells(step),offset:step/stages,easing:'steps(1,end)'})),'linear');
  }else if(effect==='randomBar'){
    const bands=16,vertical=options.direction==='vert';
    const ranks=Array.from({length:bands},(_,i)=>i);
    for(let i=bands-1;i>0;i--){
      const j=Math.floor(Math.random()*(i+1));
      [ranks[i],ranks[j]]=[ranks[j],ranks[i]];
    }
    const bars=(step)=>{
      const points=[];
      for(let band=0;band<bands;band++){
        const top=100*band/bands,bottom=top+100/bands*Math.max(0,Math.min(1,step-ranks[band]));
        // Retrace the edge bridge so disconnected bars do not clip diagonal slivers.
        for(const [x,y] of [[0,0],[0,top],[100,top],[100,bottom],[0,bottom],[0,top],[0,0]])
          points.push((vertical?y:x)+'% '+(vertical?x:y)+'%');
      }
      return 'polygon('+points.join(',')+')';
    };
    animate(incoming,Array.from({length:bands+1},(_,step)=>({clipPath:bars(step),offset:step/bands})));
  }else if(effect==='strips'){
    const direction=options.direction??'lu';
    const strips=(progress)=>{
      const points=[];
      for(let band=0;band<8;band++){
        const top=band*12.5,bottom=(band+1)*12.5;
        const right=100*Math.max(0,Math.min(1,(15*progress-band)/8));
        points.push([0,0],[0,top],[right,top],[right,bottom],[0,bottom],[0,top],[0,0]);
      }
      return 'polygon('+points.map(([x,y])=>(direction.startsWith('l')?100-x:x)+'% '+(direction.endsWith('u')?100-y:y)+'%').join(',')+')';
    };
    // Include every band's start/end so interpolation preserves the staggered wipe.
    animate(incoming,Array.from({length:16},(_,step)=>({clipPath:strips(step/15),offset:step/15})));
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
