import { Terminal } from '@xterm/xterm';
import { FitAddon } from '@xterm/addon-fit';

const element = (id: string) => document.getElementById(id)!;
const host = element('terminal');
const start = element('terminal-start') as HTMLButtonElement;
const stop = element('terminal-stop') as HTMLButtonElement;
const status = element('terminal-status');
const key = 'office-kit-terminal-client';
const client = sessionStorage.getItem(key) || crypto.randomUUID();
sessionStorage.setItem(key, client);
const terminal = new Terminal({
  fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Consolas, monospace',
  fontSize: 12,
  lineHeight: 1.2,
  cursorBlink: true,
  scrollback: 5000,
  theme: {
    background: '#171b24',
    foreground: '#e4e8f0',
    cursor: '#aebdff',
    selectionBackground: '#4b5980',
  },
});
const fit = new FitAddon();
terminal.loadAddon(fit);
terminal.open(host);
let running = false,
  owned = false;
let queue = Promise.resolve();
const focus = () =>
  JSON.parse(element('chat-context').dataset.focus || '{"slide":null,"revision":0}');
async function action(path: string, data: Record<string, unknown>) {
  const response = await fetch('/terminal/' + path, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ client, ...data }),
  });
  const result = await response.json();
  if (!response.ok) throw new Error(result.error);
}
function send(path: string, data: Record<string, unknown>) {
  // Preserve keystroke order even while earlier requests are in flight.
  queue = queue
    .then(() => action(path, data))
    .catch((error: Error) => {
      status.textContent = error.message;
    });
  return queue;
}
function resize() {
  if (!host.clientWidth || !host.clientHeight) return;
  fit.fit();
  if (running && owned) void send('resize', { cols: terminal.cols, rows: terminal.rows });
}
new ResizeObserver(resize).observe(host);
terminal.onData((data) => {
  if (running && owned) void send('input', { data, ...focus() });
});
// Navigation belongs to the terminal while it has focus, including its menus.
host.addEventListener('keydown', (event) => event.stopPropagation());
start.onclick = async () => {
  start.disabled = true;
  resize();
  await send('start', { cols: terminal.cols, rows: terminal.rows, ...focus() });
  start.disabled = running;
  terminal.focus();
};
stop.onclick = () => {
  void send('stop', {});
};
const events = new EventSource('/terminal/events');
events.addEventListener('replay', (event) => {
  terminal.reset();
  terminal.write(JSON.parse(event.data));
});
events.addEventListener('output', (event) => terminal.write(JSON.parse(event.data)));
events.addEventListener('state', (event) => {
  const state = JSON.parse(event.data);
  const wasRunning = running;
  running = state.running;
  owned = state.owner === client;
  terminal.options.disableStdin = !running || !owned;
  start.hidden = running;
  start.disabled = running;
  stop.hidden = !running || !owned;
  status.textContent = running && !owned ? 'Open in another tab · view only' : state.status;
  element('chat-provider').toggleAttribute('disabled', running);
  if (running && !wasRunning) resize();
});
events.onerror = () => {
  status.textContent = 'Reconnecting to Claude Code…';
};
