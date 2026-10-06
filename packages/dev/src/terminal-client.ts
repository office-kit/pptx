import { mountWorkspace } from './workspace-client.ts';
import { Terminal } from '@xterm/xterm';
import { FitAddon } from '@xterm/addon-fit';

// ANSI colors tuned for contrast on the editor's pane background in each scheme.
const ANSI_LIGHT = {
  black: '#15171c',
  red: '#c92a2a',
  green: '#2b8a3e',
  yellow: '#a35d00',
  blue: '#1c5fc0',
  magenta: '#b4235a',
  cyan: '#0b7285',
  white: '#8a909c',
  brightBlack: '#545b69',
  brightRed: '#e03131',
  brightGreen: '#37a14a',
  brightYellow: '#c27400',
  brightBlue: '#2f74d6',
  brightMagenta: '#d6336c',
  brightCyan: '#1098ad',
  brightWhite: '#c9cdd6',
};
const ANSI_DARK = {
  black: '#1f1f1f',
  red: '#f1707a',
  green: '#7bc88a',
  yellow: '#e8c06a',
  blue: '#7aaef5',
  magenta: '#e85a90',
  cyan: '#6fcfdc',
  white: '#c4c4c4',
  brightBlack: '#8f8f8f',
  brightRed: '#ff9aa2',
  brightGreen: '#9fe0ab',
  brightYellow: '#f5d78e',
  brightBlue: '#a3c8ff',
  brightMagenta: '#ff9cbf',
  brightCyan: '#9ae3ec',
  brightWhite: '#f0f0f0',
};
const SELECTION_ALPHA = '4d';
const darkScheme = matchMedia('(prefers-color-scheme: dark)');
const token = (name: string) =>
  getComputedStyle(document.documentElement).getPropertyValue(name).trim();
function terminalTheme() {
  return {
    ...(darkScheme.matches ? ANSI_DARK : ANSI_LIGHT),
    background: token('--ok-panel'),
    foreground: token('--ok-text'),
    cursor: token('--ok-accent'),
    cursorAccent: token('--ok-panel'),
    // Translucent accent so selected text keeps its own ANSI color.
    selectionBackground: token('--ok-accent') + SELECTION_ALPHA,
  };
}

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
    fontFamily: token('--ok-mono'),
    fontSize: 12,
    lineHeight: 1.2,
    cursorBlink: true,
    scrollback: 5000,
    theme: terminalTheme(),
  });
  // The pane follows the system appearance like the editor; xterm paints its
  // own canvas, so it is told when the scheme flips.
  darkScheme.addEventListener('change', () => {
    terminal.options.theme = terminalTheme();
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
  window.addEventListener('message', (event) => {
    if (
      event.origin !== location.origin ||
      event.source !== parent ||
      event.data?.type !== 'inline-edit' ||
      (element('chat-provider') as HTMLSelectElement).value !== 'claude'
    )
      return;
    void (async () => {
      try {
        if (!running || !owned) throw new Error('Start Claude Code in this agent pane first.');
        const prompt = queue.then(() =>
          action('prompt', { message: event.data.message, ...event.data.focus }),
        );
        queue = prompt.catch(() => undefined);
        await prompt;
        parent.postMessage({ type: 'inline-result', id: event.data.id }, location.origin);
      } catch (error) {
        parent.postMessage(
          {
            type: 'inline-result',
            id: event.data.id,
            error: error instanceof Error ? error.message : String(error),
          },
          location.origin,
        );
      }
    })();
  });
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
    if (
      event.key === 'Enter' &&
      event.shiftKey &&
      !event.ctrlKey &&
      !event.altKey &&
      !event.metaKey &&
      !event.isComposing
    ) {
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
