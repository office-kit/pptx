/** Separate window so speaker notes never appear on the audience surface. */
export const presenterPage = `<!doctype html>
<html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>Presenter view — Office Kit</title>
<style>
*{box-sizing:border-box}body{margin:0;background:#141929;color:#e4e8f0;font:15px system-ui,sans-serif;height:100vh;display:flex;flex-direction:column}header,footer{display:flex;align-items:center;gap:14px;padding:16px;flex-wrap:wrap}h1{font-size:18px;margin:0;flex:1}h2{font-size:13px;color:#b9c2d8;margin:0 0 10px}button{font:inherit;color:inherit;background:#29243e;border:1px solid #675487;border-radius:6px;padding:8px 12px;cursor:pointer}button:disabled{opacity:.4;cursor:default}button:focus-visible{outline:2px solid #bca3ff;outline-offset:3px}main{display:grid;grid-template-columns:minmax(0,3fr) minmax(240px,2fr);gap:20px;padding:0 20px;flex:1;min-height:0;overflow:auto}section{min-width:0}#current,#next{background:#090c14;aspect-ratio:16/9}#current svg,#next svg{width:100%;height:100%}#notes{white-space:pre-wrap;overflow-wrap:anywhere;line-height:1.6;max-height:40vh;overflow:auto;padding:14px;background:#1e2538;border-radius:8px}#notes-title{margin-top:20px}#count{flex:1}#message{padding:0 20px;color:#ffc9bd}#elapsed{font-variant-numeric:tabular-nums}footer{justify-content:flex-end}@media(max-width:640px){main{grid-template-columns:1fr}#notes{max-height:none}}
</style>
<header><h1 id="title">Presenter view</h1><span id="timer-label">Elapsed</span><output id="elapsed">00:00</output><button id="reset">Reset timer</button></header>
<p id="message" role="status">Connecting to presentation…</p>
<main><section><h2 id="current-title">Current slide</h2><div id="current"></div><h2 id="notes-title">Speaker notes</h2><div id="notes" tabindex="0"></div></section><section><h2 id="next-title">Next slide</h2><div id="next"></div><p id="end" hidden></p></section></main>
<footer><span id="count" aria-live="polite"></span><button id="prev" disabled>Previous</button><button id="next-button" disabled>Next</button><button id="exit">Exit presentation</button></footer>
<script>
const byId=id=>document.getElementById(id);
const canvases={current:byId('current').attachShadow({mode:'open'}),next:byId('next').attachShadow({mode:'open'})};
const displayed={};
let locale='en',started=performance.now(),connected=false;
const text=(en,ja)=>locale==='ja'?ja:en;
// The presenter shows the slide the audience is looking at, so it runs the very
// same player over its own copy, resumed at the same point in the same stop,
// rather than being told after the fact what is on screen.
let makePlayer,animationPlayer,animationKey,animationGeneration,lastState;
const presenterReducedMotion=matchMedia('(prefers-reduced-motion: reduce)');
import('/animation-player.js').then(module=>{makePlayer=module.createAnimationPlayer;if(lastState)syncAnimation(lastState);},()=>{});
function syncAnimation(data){
 const steps=data.presenting?data.animationSteps:null;
 if(!steps||!steps.length||!makePlayer){
  if(animationPlayer){animationPlayer.dispose();animationPlayer=undefined;animationKey=undefined;animationGeneration=undefined;}
  return;
 }
 const key=data.index+'\\u0000'+(data.current??'')+'\\u0000'+JSON.stringify(steps);
 if(key!==animationKey){
  animationPlayer?.dispose();
  animationPlayer=makePlayer({root:canvases.current,steps,reducedMotion:()=>presenterReducedMotion.matches});
  animationKey=key;animationGeneration=undefined;
 }
 const progress=data.animation;
 if(!progress)return;
 // Follow the audience only when it actually moved. The effects inside a stop
 // run on the same clock here, so re-running them on every passing report would
 // restart a fade; a seek that lands on the cursor it is already on — jumping
 // to the end of the stop in hand, say — still has to be followed, which is
 // what the generation says and the cursor alone does not.
 if(progress.generation===animationGeneration)return;
 animationGeneration=progress.generation;
 animationPlayer.resume(progress.cursor,progress.elapsed);
}
function send(action,index){if(window.opener&&!window.opener.closed)window.opener.postMessage({type:'presenter-command',action,index},location.origin);}
for(const canvas of Object.values(canvases))canvas.addEventListener('click',event=>{
 const link=event.composedPath().find(node=>node instanceof Element&&node.localName==='a');
 const href=link?.getAttribute('href')??link?.getAttributeNS('http://www.w3.org/1999/xlink','href')??'';
 if(href.startsWith('#slide-')){
  event.preventDefault();
  const number=Number(href.slice(7));
  if(Number.isInteger(number)&&number>=1)send('jump',number-1);
 }
});
function update(data){
 connected=true;locale=data.locale;document.documentElement.lang=locale;
 for(const [id,en,ja] of [['title','Presenter view','発表者ビュー'],['timer-label','Elapsed','経過時間'],['reset','Reset timer','タイマーをリセット'],['current-title','Current slide','現在のスライド'],['notes-title','Speaker notes','発表者ノート'],['next-title','Next slide','次のスライド'],['prev','Previous','前へ'],['next-button','Next','次へ'],['exit','Exit presentation','プレゼンテーションを終了']])byId(id).textContent=text(en,ja);
 document.title=text('Presenter view — Office Kit','発表者ビュー — Office Kit');
 byId('message').textContent=data.presenting?'':text('Presentation is stopped.','プレゼンテーションは停止中です。');
 const animation=data.animation&&data.animation.stops>0
  ? ' · '+text('Click '+data.animation.cursor+' of '+data.animation.stops,'クリック '+data.animation.cursor+' / '+data.animation.stops)
  : '';
 byId('count').textContent=data.count?text('Slide '+(data.index+1)+' of '+data.count,'スライド '+(data.index+1)+' / '+data.count)+animation:text('No slides','スライドがありません');
 for(const name of ['current','next']){
  const svg=data[name];
  if(displayed[name]!==svg){displayed[name]=svg;canvases[name].innerHTML=svg?'<style>svg{display:block;width:100%;height:100%}</style>'+svg:'';}
  byId(name).style.aspectRatio=String(data.aspectRatio);byId(name).hidden=!svg;
 }
 byId('notes').textContent=data.notes||text('No speaker notes.','発表者ノートはありません。');
 byId('prev').disabled=!data.hasPrevious;
 // The next click may be an effect on this slide rather than the next slide.
 byId('next-button').disabled=!(data.hasNext??data.next);
 byId('exit').disabled=!data.presenting;
 byId('end').hidden=!!(data.hasNext??data.next);byId('end').textContent=text('End of presentation','プレゼンテーションの最後です');
 lastState=data;syncAnimation(data);
}
window.addEventListener('message',event=>{if(event.origin===location.origin&&event.source===window.opener&&event.data?.type==='presenter-state')update(event.data);});
byId('prev').onclick=()=>send('previous');byId('next-button').onclick=()=>send('next');byId('exit').onclick=()=>send('exit');
byId('reset').onclick=()=>{started=performance.now();byId('elapsed').textContent='00:00';};
document.addEventListener('keydown',event=>{
 if(event.altKey||event.ctrlKey||event.metaKey||event.target.closest('button'))return;
 const action=['ArrowRight','ArrowDown','PageDown',' '].includes(event.key)?'next':['ArrowLeft','ArrowUp','PageUp'].includes(event.key)?'previous':event.key==='Escape'?'exit':null;
 if(action){event.preventDefault();send(action);}
});
setInterval(()=>{
 const seconds=Math.floor((performance.now()-started)/1000);byId('elapsed').textContent=String(Math.floor(seconds/60)).padStart(2,'0')+':'+String(seconds%60).padStart(2,'0');
 if(!window.opener||window.opener.closed){byId('message').textContent=text('The presentation window is closed.','プレゼンテーションのウィンドウは閉じられています。');for(const id of ['prev','next-button','exit'])byId(id).disabled=true;}
 else if(!connected)send('ready');
},250);
send('ready');
</script></html>`;
