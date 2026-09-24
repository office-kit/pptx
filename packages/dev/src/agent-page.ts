import { previewStyles } from './styles.ts';
export const agentPage = `<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width"><link rel="stylesheet" href="/terminal.css"><style>${previewStyles}
body{display:flex;flex-direction:column;background:#191e30;height:100dvh;min-width:0;overflow:auto;overflow-x:hidden} .chat-heading{flex-wrap:wrap} .chat-heading select{min-width:100px}.chat-heading,#chat-context{flex-shrink:0}#claude-panel{min-height:100px}#codex-panel{min-height:240px}

body{background:#252525;color:#e5e5e5;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}
button,select{border-radius:3px;background:#333;border-color:#555}button:hover:not(:disabled){background:#454545}
.chat-heading{background:#292929;border-color:#444}.chat-heading select{color:#eee}.chat-heading option{background:#292929}
#chat-context{background:#252525;color:#aaa;border-color:#444}#claude-panel{background:#252525}#terminal-start,#chat-send{background:#a63c24;border-color:#bd5139;box-shadow:none;color:white}
#chat-input{background:#202020;border-color:#555;border-radius:3px;color:#eee}#chat-input:focus{border-color:#e98b73;box-shadow:none}#chat-form{border-color:#444}.chat-message.user{background:#383838;border-color:#555;border-radius:4px}.chat-message small,.chat-intro,#chat-status,.chat-shortcut{color:#aaa}.markdown pre,.markdown code{background:#303030;color:#eee;border-color:#555}.markdown a{color:#efab98}
</style><div class="chat-heading"><select id="chat-provider" aria-label="AI provider"><option value="claude">Claude Code</option><option value="codex">Codex</option></select><button id="chat-reset" hidden>New chat</button></div>
<div id="chat-context">No slide selected</div>
<section id="claude-panel" aria-label="Claude Code terminal"><div class="terminal-toolbar"><span id="terminal-status" role="status">Start Claude Code to edit your slides.</span><button id="terminal-start">Start</button><button id="terminal-stop" hidden>End session</button></div><div class="terminal-frame"><div id="terminal"></div></div></section>
<div id="codex-panel" hidden>
<div id="messages" role="log" aria-label="Conversation"><p class="chat-intro">Describe what you want to change. The current slide is included automatically. You can also ask about other slides or the whole deck.<br><br>Uses your locally installed Codex. Sign in through its CLI first. Sending a message allows it to edit this project.</p></div>
<form id="chat-form"><textarea id="chat-input" aria-label="Message" maxlength="16000" placeholder="Ask for a change…" required></textarea><div class="chat-actions"><span class="chat-shortcut">⌘ / Ctrl + Enter</span><button id="chat-stop" type="button" hidden>Stop</button><button id="chat-send" type="submit">Send</button></div></form>
<div id="chat-status" role="status">Ready</div>
</div>
<script>
const byId=id=>document.getElementById(id),base=location.pathname;
let focus={slide:null,revision:0};
window.addEventListener('message',event=>{
 if(event.origin!==location.origin||event.source!==parent)return;
 if(event.data?.type==='focus'){focus=event.data.focus;byId('chat-context').dataset.focus=JSON.stringify(focus);byId('chat-context').textContent=event.data.label;}
 if(event.data?.type==='chat')void refreshChat();
});
parent.postMessage({type:'agent-ready'},location.origin);
let chatRequest=0,chatBusy=false,lastMessages='';
function renderChat(chat){
  chatBusy=chat.running;
  byId('chat-send').disabled=chat.running;
  byId('chat-reset').disabled=chat.running;
  if(byId('chat-provider').value==='codex')byId('chat-provider').disabled=chat.running;
  byId('chat-stop').hidden=!chat.running;
  byId('chat-status').textContent=chat.status;

  const serialized=JSON.stringify(chat.messages);
  if(serialized===lastMessages)return;
  lastMessages=serialized;
  const log=byId('messages'),atBottom=log.scrollHeight-log.scrollTop-log.clientHeight<60;
  log.replaceChildren();
  if(!chat.messages.length){const intro=document.createElement('p');intro.className='chat-intro';intro.textContent='Ask for a focused slide edit or a change across the deck. Uses your locally installed CLI and its login. Sending allows edits to this project.';log.append(intro);}
  for(const message of chat.messages){
    const item=document.createElement('div'),label=document.createElement('small'),body=document.createElement('div');
    item.className='chat-message '+message.role;
    label.textContent=message.role==='user'?'You · '+(message.context?.slide?'Slide '+message.context.slide:'Whole project'):'Assistant';
    if(message.role==='assistant'&&message.html){body.className='markdown';body.innerHTML=message.html;}else body.textContent=message.text||(chat.running?'Working…':'No response');item.append(label,body);log.append(item);
  }
  if(atBottom)log.scrollTop=log.scrollHeight;
}
async function refreshChat(){
  const id=++chatRequest;
  try{const response=await fetch(base+'/chat');if(!response.ok)throw new Error('Chat unavailable');const chat=await response.json();if(id===chatRequest)renderChat(chat);}
  catch(error){if(id===chatRequest)byId('chat-status').textContent=error.message;}
}
async function chatAction(path,payload){
  ++chatRequest;
  const response=await fetch(base+path,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
  const result=await response.json();if(!response.ok)throw new Error(result.error||'Chat request failed');
  await refreshChat();
}
byId('chat-form').onsubmit=async event=>{
  event.preventDefault();if(chatBusy)return;
  const input=byId('chat-input'),message=input.value.trim();if(!message)return;
  chatBusy=true;byId('chat-send').disabled=true;
  try{await chatAction('/chat',{message,provider:byId('chat-provider').value,...focus});if(input.value.trim()===message)input.value='';}
  catch(error){byId('chat-status').textContent=error.message;chatBusy=false;byId('chat-send').disabled=false;}
};
byId('chat-input').onkeydown=event=>{if(event.key==='Enter'&&(event.metaKey||event.ctrlKey)&&!event.isComposing){event.preventDefault();byId('chat-form').requestSubmit();}};
for(const action of ['stop','reset'])byId('chat-'+action).onclick=()=>chatAction('/chat/'+action,{}).catch(error=>{byId('chat-status').textContent=error.message;});
byId('chat-provider').onchange=()=>{try{sessionStorage.setItem('office-kit-provider'+base,byId('chat-provider').value);}catch{}const claude=byId('chat-provider').value==='claude';byId('claude-panel').hidden=!claude;byId('codex-panel').hidden=claude;byId('chat-reset').hidden=claude;};
try{byId('chat-provider').value=sessionStorage.getItem('office-kit-provider'+base)||'claude';}catch{}
byId('chat-provider').onchange();
void refreshChat();
</script><script type="module" src="/terminal.js"></script></html>`;
