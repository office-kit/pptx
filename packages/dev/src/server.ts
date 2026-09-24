import {
  getSlides,
  loadPresentation,
  savePresentation,
  validatePresentation,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '@office-kit/pptx-preview';
import {
  editorModel,
  restoreEditorSession,
  readHistory,
  writeHistory,
  recordEdit,
  emptyHistory,
  editPath,
  type EditCommand,
} from './editor.ts';
import { renderPreview } from './preview-cache.ts';
import { createServer, type ServerResponse } from 'node:http';
import { watch } from 'node:fs';
import { dirname, resolve, sep } from 'node:path';
import type { BuildResult } from './build.ts';
import { createDeckBuilder } from './build-runner.ts';
import { page } from './page.ts';
import { agentPage } from './agent-page.ts';
import { readFile } from 'node:fs/promises';
import { createTerminal } from './terminal.ts';
import { createChat } from './chat.ts';

export async function serveDeck(entry: string, port = 4173) {
  let latest: BuildResult | undefined;
  let error: string | null = null;
  let building = false;
  let revision = 0;
  let generation = 0;
  let patch: Record<number, string> = {};
  let pending = false;
  let closed = false;
  let editing = false;
  let historyState = emptyHistory();
  let writingHistory: string | null = null;
  async function persistHistory(history: typeof historyState) {
    writingHistory = JSON.stringify(history);
    try {
      await writeHistory(entry, history);
      historyState = history;
    } finally {
      writingHistory = null;
    }
  }
  const clients = new Set<ServerResponse>();
  const sessions = new Map<
    string,
    {
      chat: ReturnType<typeof createChat>;
      terminal: ReturnType<typeof createTerminal>;
      closing: boolean;
    }
  >();
  function session(id: string) {
    let result = sessions.get(id);
    if (!result) {
      const chat = createChat(
        entry,
        () => {
          for (const client of clients) client.write('data: chat\n\n');
        },
        () => terminal.isRunning(),
      );
      const terminal = createTerminal(
        entry,
        () => chat.isRunning(),
        id ? '/agents/' + id : '',
        verify,
      );
      result = { chat, terminal, closing: false };
      sessions.set(id, result);
    }
    return result;
  }
  const server = createServer((request, response) => {
    const host = request.headers.host;
    if (host !== `127.0.0.1:${actualPort}` && host !== `localhost:${actualPort}`) {
      response.writeHead(403).end();
      return;
    }
    response.setHeader('Cache-Control', 'no-store');
    const match = request.url?.match(
      /^\/agents\/([\w-]{1,64})(\/(?:chat(?:\/(?:stop|reset))?|terminal\/(?:start|input|resize|stop|events|context|verify)|close))?$/,
    );
    if (request.url?.startsWith('/agents/') && !match) {
      response.writeHead(404).end();
      return;
    }
    const id = match?.[1] ?? '';
    if (match) {
      if (
        request.headers['sec-fetch-site'] === 'cross-site' ||
        (request.headers.origin && request.headers.origin !== `http://${host}`)
      ) {
        response.writeHead(403).end();
        return;
      }
      if (sessions.get(id)?.closing) {
        response.writeHead(409).end('{"error":"Agent is closing"}');
        return;
      }
      if (match[2] === '/close') {
        if (
          request.method !== 'POST' ||
          request.headers.origin !== `http://${host}` ||
          request.headers['content-type'] !== 'application/json'
        ) {
          response.writeHead(403).end();
          return;
        }
        request.resume();
        const current = sessions.get(id);
        void (async () => {
          if (current) {
            current.closing = true;
            await Promise.all([current.terminal.close(), current.chat.close()]);
            sessions.delete(id);
          }
          response.writeHead(200).end('{}');
        })().catch(() => response.writeHead(500).end());
        return;
      }
      if (!sessions.has(id) && sessions.size >= 8) {
        response.writeHead(429).end('Too many agent sessions. Close an existing pane.');
        return;
      }
      if (!match[2]) {
        if (request.method !== 'GET') {
          response.writeHead(405).end();
          return;
        }
        session(id);
        response.writeHead(200, { 'Content-Type': 'text/html' }).end(agentPage);
        return;
      }
      if (!sessions.has(id)) {
        response.writeHead(404).end();
        return;
      }
      request.url = match[2];
    }

    if (request.url === '/edit') {
      if (
        request.method !== 'POST' ||
        request.headers.origin !== `http://${host}` ||
        request.headers['sec-fetch-site'] === 'cross-site' ||
        request.headers['content-type'] !== 'application/json'
      ) {
        response.writeHead(403).end();
        return;
      }
      if (editing || building || timer || pending) {
        response
          .writeHead(409)
          .end(JSON.stringify({ error: 'The presentation is updating. Try again.' }));
        return;
      }
      editing = true;
      void (async () => {
        let body = '';
        for await (const chunk of request) {
          body += String(chunk);
          if (body.length > 28_000_000) throw new Error('Edit is too large.');
        }
        const input = JSON.parse(body) as {
          revision: number;
          command?: EditCommand;
          action?: 'undo' | 'redo';
          preview?: boolean;
          snapshot?: boolean;
          previewRender?: boolean;
        };
        if (input.revision !== revision || building || timer || pending)
          throw new Error('The presentation changed. Review it and try again.');
        const started = generation;
        const history = await readHistory(entry);
        if (JSON.stringify(history) !== JSON.stringify(historyState))
          throw new Error('Edit history changed. Reload and try again.');
        if (input.snapshot && (input.command || input.action || input.preview))
          throw new Error('Cannot combine snapshot and edit operations.');
        if (input.preview && input.action) throw new Error('Cannot preview history actions.');
        if (input.action) {
          if (input.action === 'undo' && history.cursor > 0) history.cursor--;
          else if (input.action === 'redo' && history.cursor < history.entries.length)
            history.cursor++;
          else throw new Error('Nothing to ' + input.action + '.');
          if (generation !== started) throw new Error('Source changed while editing. Try again.');
          await persistHistory(history);
          if (latest) latest = { ...latest, history };
          await rebuild();
          response.writeHead(200, { 'Content-Type': 'application/json' }).end('{}');
          return;
        }
        if (error || !latest) throw new Error('Resolve the build error before editing.');
        if (input.snapshot) {
          response
            .writeHead(200, { 'Content-Type': 'application/json' })
            .end(
              JSON.stringify({ snapshot: Buffer.from(latest.bytes).toString('base64'), revision }),
            );
          return;
        }
        if (!input.command) throw new Error('Missing edit command.');
        const presentation = await loadPresentation(latest.bytes);
        restoreEditorSession(presentation, latest.editor);
        const record = recordEdit(presentation, input.command);
        const diagnostics = validatePresentation(presentation);
        if (diagnostics.some((issue) => issue.severity === 'error'))
          throw new Error('The edit would produce an invalid presentation.');
        if (input.preview) {
          const svg = input.previewRender
            ? renderSlideToSvg(presentation, getSlides(presentation)[input.command.slide]!)
            : undefined;
          if (generation !== started) throw new Error('Source changed while editing. Try again.');
          response
            .writeHead(200, { 'Content-Type': 'application/json' })
            .end(JSON.stringify({ editor: editorModel(presentation), revision, svg }));
          return;
        }
        const bytes = await savePresentation(presentation);
        const saved = await loadPresentation(bytes);
        restoreEditorSession(saved, editorModel(presentation));
        const preview = renderPreview(saved, bytes);
        const model = editorModel(saved);
        history.entries = history.entries.slice(0, history.cursor).concat(record);
        history.cursor++;
        if (generation !== started) throw new Error('Source changed while editing. Try again.');
        await persistHistory(history);
        // Publish only after the atomic history write succeeds, then verify replay
        // before acknowledging the edit. Filesystem notifications may arrive later.
        patch = {};
        preview.slides.forEach((svg, index) => {
          if (latest?.slides[index] !== svg) patch[index] = svg;
        });
        latest = {
          ...latest,
          bytes,
          slides: preview.slides,
          slideTexts: preview.slideTexts,
          editor: model,
          history,
          diagnostics,
        };
        revision++;
        await rebuild();
        response
          .writeHead(200, { 'Content-Type': 'application/json' })
          .end(JSON.stringify({ revision }));
      })()
        .catch((cause) => {
          if (!response.headersSent)
            response
              .writeHead(409, { 'Content-Type': 'application/json' })
              .end(
                JSON.stringify({ error: cause instanceof Error ? cause.message : String(cause) }),
              );
        })
        .finally(() => {
          editing = false;
        });
    } else if (
      request.url === '/chat' ||
      request.url?.startsWith('/chat/') ||
      request.url?.startsWith('/terminal/')
    ) {
      const current = session(id);
      const handler = request.url.startsWith('/terminal/') ? current.terminal : current.chat;
      void handler.handle(request, response, (index, viewedRevision) => {
        if (viewedRevision !== revision)
          throw new Error('Preview changed. Review the current slide and send again.');
        if (index !== null && (!latest || index >= latest.slides.length))
          throw new Error('Slide no longer exists.');
        return {
          slide: index === null ? null : index + 1,
          count: latest?.slides.length ?? 0,
          revision,
          text: index === null ? '' : (latest?.slideTexts[index] ?? '').slice(0, 12000),
          entry: resolve(entry),
          files: latest?.dependencies ?? [resolve(entry)],
          buildError: error,
        };
      });
    } else if (
      request.method === 'GET' &&
      ['/terminal.js', '/terminal.css', '/editor.js'].includes(request.url ?? '')
    ) {
      const asset =
        request.url === '/editor.js'
          ? 'editor-client.js'
          : request.url === '/terminal.js'
            ? 'terminal-client.js'
            : 'terminal-client.css';
      void readFile(new URL(asset, import.meta.url)).then(
        (bytes) => {
          response.writeHead(200, {
            'Content-Type': asset.endsWith('.css') ? 'text/css' : 'text/javascript',
          });
          response.end(bytes);
        },
        () => response.writeHead(500).end('Terminal assets unavailable. Rebuild pptx-dev.'),
      );
    } else if (request.method !== 'GET') {
      response.writeHead(405).end();
    } else if (request.url === '/events') {
      response.writeHead(200, { 'Content-Type': 'text/event-stream' });
      response.write('data: ready\n\n');
      clients.add(response);
      request.on('close', () => clients.delete(response));
    } else if (request.url === '/state' || request.url?.startsWith('/state?')) {
      const since = new URL(request.url, 'http://localhost').searchParams.get('since');
      const incremental =
        since !== null && (since === String(revision) || since === String(revision - 1));
      response.writeHead(200, { 'Content-Type': 'application/json' });
      response.end(
        JSON.stringify({
          revision,
          building: building || !!timer || pending,
          ...(incremental
            ? {
                changes: since === String(revision) ? {} : patch,
                count: latest?.slides.length ?? 0,
              }
            : { slides: latest?.slides ?? [] }),
          editor: latest?.editor,
          history: {
            undo: historyState.cursor,
            redo: historyState.entries.length - historyState.cursor,
          },
          aspectRatio: latest?.aspectRatio ?? 16 / 9,
          error,
          diagnostics: latest?.diagnostics ?? [],
        }),
      );
    } else if (request.url === '/deck.pptx' && latest) {
      response.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': 'attachment; filename="deck.pptx"',
      });
      response.end(latest.bytes);
    } else if (request.url === '/') {
      response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      response.end(page);
    } else {
      response.writeHead(404).end();
    }
  });
  async function verify() {
    // A fresh build also covers writes whose watcher notification has not arrived yet.
    void rebuild();
    const deadline = Date.now() + 30000;
    while (!closed && (building || timer || pending)) {
      if (Date.now() >= deadline)
        return 'Preview verification timed out. Inspect the build before claiming success.';
      await new Promise((done) => setTimeout(done, 50));
    }
    return closed ? 'Preview server closed.' : error;
  }
  async function rebuild() {
    if (closed) return;
    if (building) {
      pending = true;
      return;
    }
    building = true;
    const started = generation;
    for (const client of clients) client.write('data: building\n\n');
    try {
      const result = await builder.build();
      if (started === generation && !closed) {
        patch = {};
        result.slides.forEach((svg, index) => {
          if (svg !== latest?.slides[index]) patch[index] = svg;
        });
        latest = result;
        historyState = result.history;
        revision++;
        error = null;
      }
    } catch (cause) {
      try {
        historyState = await readHistory(entry);
      } catch {
        /* Preserve the last readable history. */
      }
      if (started === generation)
        error = cause instanceof Error ? (cause.stack ?? cause.message) : String(cause);
    } finally {
      building = false;
      for (const client of clients) client.write('data: updated\n\n');
      if (pending && !timer) {
        pending = false;
        void rebuild();
      }
    }
  }
  let timer: ReturnType<typeof setTimeout> | undefined;
  const root = dirname(resolve(entry));
  const watcher = watch(root, { recursive: true }, async (_, filename) => {
    if (
      !filename ||
      filename
        .split(sep)
        .some((part) => ['node_modules', '.git', 'dist', '.office-kit'].includes(part))
    )
      return;
    if (!/\.([cm]?[jt]sx?|json|pptx|png|jpe?g|gif|bmp|tiff?|emf|wmf|svg)$/i.test(filename)) return;
    if (resolve(root, filename) === resolve(editPath(entry))) {
      try {
        const history = JSON.stringify(await readHistory(entry));
        // Our own saves explicitly rebuild. A delayed notification must not
        // interrupt the user's next edit; external history changes still rebuild.
        if (history === writingHistory || history === JSON.stringify(historyState)) return;
      } catch {
        // Let the build report malformed or unreadable external history.
      }
    }
    if (closed) return;
    generation++;
    void builder.cancel();
    clearTimeout(timer);
    timer = setTimeout(() => {
      timer = undefined;
      pending = false;
      void rebuild();
    }, 30);
  });
  const builder = createDeckBuilder(entry);
  let actualPort = port;
  try {
    await new Promise<void>((resolveListen, reject) => {
      server.once('error', reject);
      server.listen(port, '127.0.0.1', () => {
        server.off('error', reject);
        resolveListen();
      });
    });
  } catch (cause) {
    watcher.close();
    await builder.close();
    throw cause;
  }
  const address = server.address();
  if (address && typeof address !== 'string') actualPort = address.port;
  await rebuild();
  return {
    url: `http://127.0.0.1:${actualPort}`,
    async close() {
      closed = true;
      clearTimeout(timer);
      watcher.close();
      await Promise.all(
        [...sessions.values()].flatMap(({ terminal, chat }) => [terminal.close(), chat.close()]),
      );
      await builder.close();
      for (const client of clients) client.end();
      await new Promise<void>((done, reject) =>
        server.close((cause) => (cause ? reject(cause) : done())),
      );
    },
  };
}
