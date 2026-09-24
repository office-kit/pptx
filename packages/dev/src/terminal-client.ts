import { mountWorkspace } from './workspace-client.ts';
import { Terminal } from '@xterm/xterm';
import { FitAddon } from '@xterm/addon-fit';

export function mountTerminal() {
  const base = location.pathname === '/' ? '' : location.pathname;
  const element = (id: string) => document.getElementById(id)!;
  const host = element('terminal');
  const start = element('terminal-start') as HTMLButtonElement;
  const stop = element('terminal-stop') as HTMLButtonElement;
  const status = element('terminal-status');
  const key = 'office-kit-terminal-client' + base;
  const client = sessionStorage.getItem(key) || crypto.randomUUID();
  sessionStorage.setItem(key, client);
  const terminal = new Terminal({
    fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Consolas, monospace',
    fontSize: 12,
    lineHeight: 1.2,
    cursorBlink: true,
    scrollback: 5000,
    theme: {
      background: '#252525',
      foreground: '#e5e5e5',
      cursor: '#eeeeee',
      selectionBackground: '#555555',
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
    const response = await fetch(base + '/terminal/' + path, {
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
    const cols = terminal.cols,
      rows = terminal.rows;
    fit.fit();
    if (running && owned && (terminal.cols !== cols || terminal.rows !== rows))
      void send('resize', { cols: terminal.cols, rows: terminal.rows });
  }
  new ResizeObserver(resize).observe(host);
  void document.fonts.ready.then(resize);
  terminal.attachCustomKeyEventHandler((event) => {
    if (event.key === 'Enter' && event.shiftKey && !event.ctrlKey && !event.altKey && !event.metaKey && !event.isComposing) {
      if (event.type === 'keydown') {
        event.preventDefault();
        if (running && owned) void send('input', { data: '\x1b[13;2u', ...focus() });
      }
      return false;
    }
    return true;
  });
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
  const events = new EventSource(base + '/terminal/events');
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
    if (running && !wasRunning) {
      resize();
      if (owned) void send('resize', { cols: terminal.cols, rows: terminal.rows });
    }
  });
  events.onerror = () => {
    status.textContent = 'Reconnecting to Claude Code…';
  };
}

if (document.getElementById('agent-workspace')) mountWorkspace();
else mountTerminal();
