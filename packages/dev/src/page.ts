export const page = `<!doctype html>
<html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>Office Kit — PowerPoint preview</title>
<style>
*{box-sizing:border-box}body{margin:0;background:#eff1f5;color:#19212d;font:14px system-ui}
header{height:60px;background:#fff;border-bottom:1px solid #dbe0e7;display:flex;align-items:center;padding:0 24px;gap:24px;position:sticky;top:0}
header strong{font-size:17px}#status{flex:1;color:#5a6575}a,button{color:inherit}button,a.download{padding:8px 14px;border:1px solid #cbd2dc;background:white;border-radius:6px;text-decoration:none;cursor:pointer}
main{max-width:1200px;margin:24px auto;padding:0 24px}#error{display:none;background:#fff0ef;color:#922e25;padding:16px;white-space:pre-wrap;border:1px solid #e8bbb7}
nav{display:flex;gap:12px;align-items:center;margin:16px 0}#slide{background:white;box-shadow:0 3px 24px #19212d18;aspect-ratio:16/9}iframe{width:100%;height:100%;border:0}footer{color:#667085;margin:20px 0}
</style>
<header><strong>Office Kit</strong><span id="status">Building…</span><a class="download" href="/deck.pptx">Download PPTX</a></header>
<main><pre id="error"></pre><nav><button id="prev">Previous</button><span id="count"></span><button id="next">Next</button></nav><div id="slide"></div><footer>Save your TSX to update. Preview rendering may differ from PowerPoint.</footer></main>
<script>
let state={slides:[],error:null},index=0;
const byId=id=>document.getElementById(id);
function render(){
  byId('status').textContent=state.error?'Build failed — showing last successful output':state.slides.length+' slides · live';
  byId('error').textContent=state.error||'';byId('error').style.display=state.error?'block':'none';
  index=Math.max(0,Math.min(index,state.slides.length-1));
  byId('count').textContent=state.slides.length?(index+1)+' / '+state.slides.length:'No slides';
  byId('prev').disabled=index===0;byId('next').disabled=index>=state.slides.length-1;
  byId('slide').style.aspectRatio=state.aspectRatio||16/9;
  byId('slide').replaceChildren();
  if(state.slides[index]){const frame=document.createElement('iframe');frame.sandbox='';frame.title='Slide '+(index+1);frame.srcdoc='<style>html,body{margin:0;width:100%;height:100%;overflow:hidden}svg{display:block;width:100%;height:100%}</style>'+state.slides[index];byId('slide').append(frame);}
}
let refreshId=0;
async function refresh(){const id=++refreshId;try{const updated=await fetch('/state').then(r=>r.json());if(id===refreshId){state=updated;render();}}catch{byId('status').textContent='Reconnecting…';}}
byId('prev').onclick=()=>{index--;render()};byId('next').onclick=()=>{index++;render()};
document.addEventListener('keydown',e=>{if(e.key==='ArrowLeft'&&index>0){index--;render()}if(e.key==='ArrowRight'&&index<state.slides.length-1){index++;render()}});
const events=new EventSource('/events');events.onmessage=refresh;events.onerror=()=>{byId('status').textContent='Reconnecting…'};refresh();
</script></html>`;
