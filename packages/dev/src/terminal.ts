import { randomBytes } from 'node:crypto';
import { dirname, resolve, join } from 'node:path';
import { createRequire } from 'node:module';
import { stat, chmod } from 'node:fs/promises';
import type { IncomingMessage, ServerResponse } from 'node:http';
import type { IPty } from 'node-pty';
import type { FocusContext } from './chat.ts';

const OUTPUT_LIMIT = 2_000_000;
const REQUEST_LIMIT = 128_000;

/** A real interactive Claude session, shared across reloads of its owning tab. */
export function createTerminal(entry: string, busy = () => false) {
  let process: IPty | undefined;
  let owner = '';
  let output = '';
  let context: FocusContext | undefined;
  let status = 'Start Claude Code to edit your slides.';
  let completion = Promise.resolve();
  let closing = false;
  let stopping = false;
  let starting = false;
  const token = randomBytes(32).toString('hex');
  const clients = new Set<ServerResponse>();
  const snapshot = () => ({ running: !!process, owner, status });
  const emit = (event: string, value: unknown) => {
    for (const client of clients) {
      if (!client.write(`event: ${event}\ndata: ${JSON.stringify(value)}\n\n`)) {
        clients.delete(client);
        client.end();
      }
    }
  };
  async function stop() {
    if (!process || stopping) return completion;
    stopping = true;
    status = 'Stopping…';
    emit('state', snapshot());
    const running = process;
    running.kill();
    const timer = setTimeout(() => {
      if (process === running) running.kill('SIGKILL');
    }, 2000);
    await completion;
    clearTimeout(timer);
  }
  async function handle(
    request: IncomingMessage,
    response: ServerResponse,
    focus: (index: number | null, revision: number) => FocusContext,
  ) {
    const json = (code: number, value: unknown) => {
      response.writeHead(code, { 'Content-Type': 'application/json' });
      response.end(JSON.stringify(value));
    };
    if (request.url === '/terminal/context') {
      if (request.method !== 'POST' || request.headers.authorization !== `Bearer ${token}`) {
        json(403, { error: 'Invalid hook credentials' });
        return;
      }
      request.resume();
      json(200, {
        hookSpecificOutput: {
          hookEventName: 'UserPromptSubmit',
          additionalContext: `You are editing the local PowerPoint TSX project at ${resolve(entry)}.
Read the project's authoring instructions. Make focused source patches and preserve unrelated work.
The preview rebuilds on save; do not start another dev server.
The focused slide is context, NOT a restriction. For "this slide", edit only its relevant source. For other slides or deck-wide requests, locate the relevant source or shared theme.
Slide numbers are 1-based. Verify source before editing; dependency paths are candidates, not an exact slide-to-file mapping.
Preview context captured with the latest terminal input: ${JSON.stringify(context)}`,
        },
      });
      return;
    }
    const origin = `http://${request.headers.host}`;
    if (
      request.headers['sec-fetch-site'] === 'cross-site' ||
      (request.headers.origin && request.headers.origin !== origin)
    ) {
      json(403, { error: 'Same-origin requests only' });
      return;
    }
    if (request.method === 'GET' && request.url === '/terminal/events') {
      response.writeHead(200, { 'Content-Type': 'text/event-stream' });
      response.write(`event: replay\ndata: ${JSON.stringify(output)}\n\n`);
      response.write(`event: state\ndata: ${JSON.stringify(snapshot())}\n\n`);
      clients.add(response);
      request.on('close', () => clients.delete(response));
      return;
    }
    if (
      request.method !== 'POST' ||
      request.headers.origin !== origin ||
      request.headers['content-type'] !== 'application/json'
    ) {
      json(403, { error: 'Same-origin JSON requests only' });
      return;
    }
    try {
      let body = '';
      request.setEncoding('utf8');
      for await (const chunk of request) {
        body += chunk;
        if (Buffer.byteLength(body) > REQUEST_LIMIT) throw new Error('Request too large');
      }
      const value = JSON.parse(body);
      if (!value || typeof value.client !== 'string' || !/^[\w-]{1,64}$/.test(value.client))
        throw new Error('Invalid terminal client');
      if (closing || starting || stopping || (process && owner !== value.client)) {
        json(409, {
          error: stopping
            ? 'Claude Code is stopping'
            : 'Claude Code is open in another tab. Use that tab to control it.',
        });
        return;
      }
      const size = () => {
        if (
          !Number.isInteger(value.cols) ||
          value.cols < 20 ||
          value.cols > 500 ||
          !Number.isInteger(value.rows) ||
          value.rows < 5 ||
          value.rows > 300
        )
          throw new Error('Invalid terminal dimensions');
      };
      const updateFocus = () => {
        if (
          !(value.slide === null || (Number.isInteger(value.slide) && value.slide >= 0)) ||
          !Number.isInteger(value.revision)
        )
          throw new Error('Invalid preview context');
        context = focus(value.slide, value.revision);
      };
      if (request.url === '/terminal/start') {
        if (busy()) throw new Error('A Codex edit is already running');
        size();
        updateFocus();
        if (!process) {
          starting = true;
          try {
            // Lazy loading leaves preview/export usable if a native PTY cannot load.
            // node-pty 1.1.0 ships its macOS spawn helper without executable bits.
            // Restore the packaged helper's mode before its first launch.
            if (globalThis.process.platform === 'darwin') {
              const root = dirname(createRequire(import.meta.url).resolve('node-pty/package.json'));
              const helper = join(
                root,
                'prebuilds',
                `darwin-${globalThis.process.arch}`,
                'spawn-helper',
              );
              const info = await stat(helper).catch((error: NodeJS.ErrnoException) => {
                if (error.code === 'ENOENT') return undefined;
                throw error;
              });
              if (info && !(info.mode & 0o100)) await chmod(helper, info.mode | 0o100);
            }
            const pty = await import('node-pty');
            if (process || closing) throw new Error('Terminal state changed. Try again.');
            const settings = {
              hooks: {
                UserPromptSubmit: [
                  {
                    hooks: [
                      {
                        type: 'http',
                        url: `${origin}/terminal/context`,
                        headers: { Authorization: `Bearer ${token}` },
                        timeout: 5,
                      },
                    ],
                  },
                ],
              },
            };
            const env = { ...globalThis.process.env };
            // This is an independent interactive session, even if the preview was started by Claude.
            delete env.CLAUDECODE;
            process = pty.spawn('claude', ['--settings', JSON.stringify(settings)], {
              cwd: dirname(resolve(entry)),
              name: 'xterm-256color',
              env,
              cols: value.cols,
              rows: value.rows,
            });
          } finally {
            starting = false;
          }
          owner = value.client;
          output = '';
          status = 'Claude Code';
          emit('replay', '');
          process.onData((data) => {
            output = (output + data).slice(-OUTPUT_LIMIT);
            emit('output', data);
          });
          completion = new Promise<void>((done) =>
            process!.onExit(({ exitCode }) => {
              process = undefined;
              stopping = false;
              status = `Claude Code exited (${exitCode}). Start to open a new session.`;
              emit('state', snapshot());
              done();
            }),
          );
        }
      } else if (request.url === '/terminal/input') {
        if (!process) throw new Error('Start Claude Code first');
        if (typeof value.data !== 'string' || value.data.length > 32000)
          throw new Error('Invalid terminal input');
        updateFocus();
        process.write(value.data);
      } else if (request.url === '/terminal/resize') {
        size();
        process?.resize(value.cols, value.rows);
      } else if (request.url === '/terminal/stop') {
        await stop();
      } else {
        json(404, { error: 'Not found' });
        return;
      }
      emit('state', snapshot());
      json(200, snapshot());
    } catch (cause) {
      json(400, { error: cause instanceof Error ? cause.message : 'Invalid terminal request' });
    }
  }
  return {
    handle,
    isRunning: () => starting || !!process,
    async close() {
      closing = true;
      await stop();
      for (const client of clients) client.end();
      clients.clear();
    },
  };
}
