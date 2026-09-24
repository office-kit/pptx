// Isolated SVG layers prevent duplicate SVG IDs from affecting the live slide.
export const transitionPlayback = `
let transitionRun=null;
let transitionAudio=null;
function stopTransitionSound(){
  const audio=transitionAudio;transitionAudio=null;
  if(!audio)return;
  audio.pause();audio.removeAttribute('src');audio.load();
}
function playTransitionSound(sound){
  if(!sound||sound.kind==='none')return;
  if(sound.kind!=='stop'&&sound.kind!=='play')return;
  stopTransitionSound();
  if(typeof stopActionSounds==='function')stopActionSounds();
  if(sound.kind==='stop'||!sound.base64)return;
  const audio=new Audio('data:audio/wav;base64,'+sound.base64);
  transitionAudio=audio;audio.loop=!!sound.loop;
  const release=()=>{if(transitionAudio===audio)stopTransitionSound();};
  audio.addEventListener('ended',release,{once:true});
  audio.addEventListener('error',release,{once:true});
  void audio.play().catch(release);
}
function cancelTransitionPlayback(){
  const run=transitionRun;transitionRun=null;
  if(!run)return;
  run.resizeObserver?.disconnect();
  for(const animation of run.animations)animation.cancel();
  run.overlay.remove();
}
function playTransition(previousSvg,settings,onComplete=()=>{}){
  cancelTransitionPlayback();
  let effect=settings?.effect??'none';
  const duration=settings?.durationMs??500;
  const supported=['fade','push','wipe','split','cover','pull','cut','zoom','circle','diamond','plus','blinds','wheel','wheelReverse','dissolve','randomBar','checker','comb','wedge','strips','newsflash'];
  if(effect==='random'){
    effect=supported[Math.floor(Math.random()*supported.length)];
    // Choose afresh for playback without replacing the document's Random setting.
    settings={effect,durationMs:duration};
  }
  if(!displayedSvg||!duration||!supported.includes(effect)||(['wheel','wheelReverse'].includes(effect)&&![1,2,3,4,8].includes(settings.spokes??4))){onComplete();return;}
  const overlay=document.createElement('div');
  overlay.dataset.transitionPlayback=effect;
  overlay.setAttribute('aria-hidden','true');
  overlay.style.cssText='position:absolute;inset:0;overflow:hidden;background:black;pointer-events:none;z-index:1000';
  const makeLayer=svg=>{
    const layer=document.createElement('div');
    layer.style.cssText='position:absolute;inset:0;background:black';
    layer.attachShadow({mode:'open'}).innerHTML='<style>svg{display:block;width:100%;height:100%}</style>'+(svg||'');
    overlay.append(layer);return layer;
  };
  const [oldLayer,newLayer]=[previousSvg,canvas.querySelector('svg')?.outerHTML||displayedSvg].map(makeLayer);
  canvas.append(overlay);
  const run={overlay,animations:[],effectRevision:0};transitionRun=run;
  const animate=(element,frames,options={})=>{
    const animation=element.animate(frames,{duration,easing:'linear',fill:'both',...options});
    run.animations.push(animation);return animation;
  };
  const dir=settings?.direction??(['split','zoom'].includes(effect)?'out':effect==='strips'?'lu':['blinds','checker','comb','randomBar'].includes(effect)?'horz':'l');
  const vectors={l:[-100,0],r:[100,0],u:[0,-100],d:[0,100],lu:[-100,-100],ru:[100,-100],ld:[-100,100],rd:[100,100]};
  const [x,y]=vectors[dir]??vectors.l;
  const translate=(a,b)=>'translate('+a+'%, '+b+'%)';
  if(effect==='fade'){
    if(settings.thruBlack){
      animate(oldLayer,[{opacity:1},{opacity:0}],{duration:duration/2});
      animate(newLayer,[{opacity:0},{opacity:1}],{delay:duration/2,duration:duration/2});
    }else animate(newLayer,[{opacity:0},{opacity:1}]);
  }else if(effect==='cut'){
    if(!settings.thruBlack){cancelTransitionPlayback();onComplete();return;}
    oldLayer.style.visibility='hidden';
    animate(newLayer,[{opacity:0},{opacity:0},{opacity:1}],{easing:'steps(1,end)'});
  }else if(['push','cover','pull'].includes(effect)){
    if(effect==='pull'){
      oldLayer.style.zIndex='1';animate(oldLayer,[{transform:translate(0,0)},{transform:translate(x,y)}]);
    }else{
      animate(newLayer,[{transform:translate(-x,-y)},{transform:translate(0,0)}]);
      if(effect==='push')animate(oldLayer,[{transform:translate(0,0)},{transform:translate(x,y)}]);
    }
  }else if(effect==='wipe'){
    const clips={l:'inset(0 0 0 100%)',r:'inset(0 100% 0 0)',u:'inset(100% 0 0 0)',d:'inset(0 0 100% 0)'};
    animate(newLayer,[{clipPath:clips[dir]??clips.l},{clipPath:'inset(0 0 0 0)'}]);
  }else if(effect==='split'){
    const center=settings.orientation==='vert'?'inset(0 50% 0 50%)':'inset(50% 0 50% 0)';
    if(dir==='in'){
      oldLayer.style.zIndex='1';animate(oldLayer,[{clipPath:'inset(0 0 0 0)'},{clipPath:center}]);
    }else animate(newLayer,[{clipPath:center},{clipPath:'inset(0 0 0 0)'}]);
  }else if(effect==='strips'){
    newLayer.remove();
    const direction=dir;
    for(let row=0;row<8;row++){
      const strip=makeLayer(displayedSvg);
      const order=direction.endsWith('u')?7-row:row;
      const clip=amount=>'inset('+(row*12.5)+'% '+(direction.startsWith('r')?100-amount:0)+'% '+((7-row)*12.5)+'% '+(direction.startsWith('l')?100-amount:0)+'%)';
      const start=order/15,end=(order+8)/15;
      const frames=[{clipPath:clip(0),offset:0}];
      if(start>0)frames.push({clipPath:clip(0),offset:start});
      frames.push({clipPath:clip(100),offset:end});
      if(end<1)frames.push({clipPath:clip(100),offset:1});
      animate(strip,frames);
    }
  }else if(effect==='comb'){
    newLayer.remove();
    for(let i=0;i<8;i++){
      const strip=makeLayer(displayedSvg);
      const before=i*12.5,after=100-(i+1)*12.5;
      strip.style.clipPath=dir==='vert'?'inset(0 '+after+'% 0 '+before+'%)':'inset('+before+'% 0 '+after+'% 0)';
      const offset=i%2?100:-100;
      animate(strip,[{transform:dir==='vert'?translate(0,offset):translate(offset,0)},{transform:translate(0,0)}]);
    }
  }else if(effect==='blinds'){
    newLayer.remove();
    const horizontal=dir!=='vert';
    for(let i=0;i<8;i++){
      const start=i*12.5,end=(i+1)*12.5;
      const strip=makeLayer(displayedSvg);
      const clip=progress=>horizontal
        ?'inset(0 '+(100-start-(end-start)*progress)+'% 0 '+start+'%)'
        :'inset('+start+'% 0 '+(100-start-(end-start)*progress)+'% 0)';
      animate(strip,[{clipPath:clip(0)},{clipPath:clip(1)}]);
    }
  }else if(['dissolve','randomBar','checker'].includes(effect)){
    const bounds=overlay.getBoundingClientRect();
    const columns=effect==='dissolve'?24:effect==='checker'?8:dir==='vert'?64:1;
    const rows=effect==='dissolve'?Math.max(1,Math.min(64,Math.round(24*bounds.height/Math.max(1,bounds.width)))):effect==='checker'?6:dir==='vert'?1:64;
    const cells=Array.from({length:columns*rows},(_,i)=>({column:i%columns,row:Math.floor(i/columns)}));
    if(effect!=='checker')for(let i=cells.length-1;i>0;i--){
      const j=Math.floor(Math.random()*(i+1));[cells[i],cells[j]]=[cells[j],cells[i]];
    }
    let width=overlay.clientWidth,height=overlay.clientHeight;
    const clip=progress=>{
      const paths=[];
      const selected=effect==='checker'?cells:cells.slice(0,Math.floor(cells.length*progress));
      for(const {column,row} of selected){
        const x=column*width/columns,y=row*height/rows;
        const amount=effect==='checker'?Math.max(0,Math.min(1,progress*2-(column+row)%2)):1;
        const w=width/columns*(effect==='checker'&&dir!=='vert'?amount:1);
        const h=height/rows*(effect==='checker'&&dir==='vert'?amount:1);
        // Independent subpaths avoid antialiased seams from polygon connecting edges.
        paths.push('M'+x+' '+y+'h'+w+'v'+h+'h'+(-w)+'Z');
      }
      return 'path("'+(paths.join(' ')||'M0 0Z')+'")';
    };
    const steps=effect==='checker'?2:32;
    const frames=()=>Array.from({length:steps+1},(_,i)=>({clipPath:clip(i/steps),offset:i/steps}));
    const animation=animate(newLayer,frames(),effect==='checker'?{}:{easing:'steps(32,end)'});
    run.resizeObserver=new ResizeObserver(()=>{
      if(width===overlay.clientWidth&&height===overlay.clientHeight)return;
      width=overlay.clientWidth;height=overlay.clientHeight;
      animation.effect.setKeyframes(frames());run.effectRevision++;
    });
    run.resizeObserver.observe(overlay);
  }else if(effect==='wheel'||effect==='wheelReverse'||effect==='wedge'){
    const spokes=effect==='wedge'?1:settings.spokes??4;
    // Fixed polygon topology lets the browser interpolate between angular samples.
    const frames=Array.from({length:61},(_,frame)=>{
      const progress=frame/60,points=[];
      for(let spoke=0;spoke<spokes;spoke++){
        points.push('50% 50%');
        for(let step=0;step<=24;step++){
          const angle=-Math.PI/2+2*Math.PI*(spoke+(effect==='wheelReverse'?-1:1)*progress*step/24)/spokes-(effect==='wedge'?Math.PI*progress:0);
          points.push((50+Math.cos(angle)*75)+'% '+(50+Math.sin(angle)*75)+'%');
        }
        points.push('50% 50%');
      }
      return {clipPath:'polygon('+points.join(',')+')',offset:progress};
    });
    animate(newLayer,frames);
  }else if(effect==='plus'){
    const cross=progress=>{
      const a=50-50*progress,b=50+50*progress;
      const c=50-150*progress,d=50+150*progress;
      return 'polygon('+[[a,c],[b,c],[b,a],[d,a],[d,b],[b,b],[b,d],[a,d],[a,b],[c,b],[c,a],[a,a]].map(([x,y])=>x+'% '+y+'%').join(',')+')';
    };
    animate(newLayer,[{clipPath:cross(0)},{clipPath:cross(1)}]);
  }else if(effect==='newsflash'){
    animate(newLayer,[{transform:'rotate(360deg) scale(0)'},{transform:'rotate(0deg) scale(1)'}]);
  }else if(effect==='zoom'){
    if(dir==='out'){oldLayer.style.zIndex='1';animate(oldLayer,[{transform:'scale(1)'},{transform:'scale(0)'}]);}
    else animate(newLayer,[{transform:'scale(0)'},{transform:'scale(1)'}]);
  }else{
    const small=effect==='circle'?'circle(0% at 50% 50%)':'polygon(50% 50%,50% 50%,50% 50%,50% 50%)';
    const large=effect==='circle'?'circle(75% at 50% 50%)':'polygon(50% -50%,150% 50%,50% 150%,-50% 50%)';
    animate(newLayer,[{clipPath:small},{clipPath:large}]);
  }
  Promise.all(run.animations.map(animation=>animation.finished)).then(()=>{
    if(transitionRun!==run)return;
    cancelTransitionPlayback();onComplete();
  }).catch(()=>{});
}
function previewTransition(){
  if(presenting||!state.slides.length)return;
  stopTransitionSound();
  playTransitionSound(state.editor?.slides[index]?.transitionSound);
  playTransition(state.slides[index-1]??'',state.editor?.slides[index]?.transition);
}
`;
