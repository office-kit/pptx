import { createTextEditor } from './text-edit.ts';
import { createVisualReviewer } from './visual-review.ts';
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
  const clients = new Set<ServerResponse>();
  const sessions = new Map<
    string,
    {
      chat: ReturnType<typeof createChat>;
      terminal: ReturnType<typeof createTerminal>;
      closing: boolean;
    }
  >();
  const capturedReviews = new Map<string, string[]>();
  function session(id: string) {
    let result = sessions.get(id);
    if (!result) {
      const review = createVisualReviewer(resolve(entry), () => latest?.slides ?? []);
      const chat = createChat(
        entry,
        () => {
          for (const client of clients) client.write('data: chat\n\n');
        },
        () => terminal.isRunning(),
        verify,
        review,
        (prompt) => capturedReviews.get(prompt) ?? [],
      );
      const terminal = createTerminal(
        entry,
        () => chat.isRunning(),
        id ? '/agents/' + id : '',
        verify,
        review,
      );
      result = { chat, terminal, closing: false };
      sessions.set(id, result);
    }
    return result;
  }
  const textEditor = createTextEditor(
    resolve(entry),
    () => {
      if (!latest || building || pending || timer || error)
        throw new Error('Wait for a successful preview build.');
      return { ...latest, revision };
    },
    verify,
  );
  const server = createServer((request, response) => {
    const host = request.headers.host;
    if (host !== `127.0.0.1:${actualPort}` && host !== `localhost:${actualPort}`) {
      response.writeHead(403).end();
      return;
    }
    response.setHeader('Cache-Control', 'no-store');
    const match = request.url?.match(
      /^\/agents\/([\w-]{1,64})(\/(?:chat(?:\/(?:stop|reset))?|terminal\/(?:start|input|prompt|resize|stop|events|context|verify)|close))?$/,
    );
    if (request.url?.startsWith('/agents/') && !match) {
      response.writeHead(404).end();
      return;
    }
    if (request.url === '/text-edit' || request.url === '/text-edit/undo') {
      if (
        request.method !== 'POST' ||
        request.headers.origin !== `http://${host}` ||
        request.headers['content-type'] !== 'application/json'
      ) {
        response.writeHead(403).end();
        return;
      }
      void (async () => {
        let body = '';
        request.setEncoding('utf8');
        for await (const chunk of request) {
          body += chunk;
          if (Buffer.byteLength(body) > 100000) throw new Error('Request too large');
        }
        const manualReview = createVisualReviewer(resolve(entry), () => latest?.slides ?? []);
        manualReview.begin();
        if (request.url === '/text-edit/undo') await textEditor.undo();
        else await textEditor.edit(JSON.parse(body));
        let review;
        let reviewError;
        try {
          review = await manualReview.next();
          if (review) {
            capturedReviews.set(review.prompt, review.images);
            if (capturedReviews.size > 64)
              capturedReviews.delete(capturedReviews.keys().next().value!);
          }
        } catch (cause) {
          reviewError = cause instanceof Error ? cause.message : String(cause);
        }
        response
          .writeHead(200, { 'Content-Type': 'application/json' })
          .end(JSON.stringify({ review, reviewError }));
      })().catch((cause) =>
        response
          .writeHead(400, { 'Content-Type': 'application/json' })
          .end(JSON.stringify({ error: cause instanceof Error ? cause.message : String(cause) })),
      );
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

    if (
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
      ['/terminal.js', '/terminal.css'].includes(request.url ?? '')
    ) {
      const asset = request.url === '/terminal.js' ? 'terminal-client.js' : 'terminal-client.css';
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
          building,
          ...(incremental
            ? {
                changes: since === String(revision) ? {} : patch,
                count: latest?.slides.length ?? 0,
              }
            : { slides: latest?.slides ?? [] }),
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
        revision++;
        error = null;
      }
    } catch (cause) {
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
  const watcher = watch(root, { recursive: true }, (_, filename) => {
    if (
      !filename ||
      filename
        .split(sep)
        .some((part) => ['node_modules', '.git', 'dist', '.office-kit'].includes(part))
    )
      return;
    if (!/\.([cm]?[jt]sx?|json|pptx|png|jpe?g|gif|bmp|tiff?|emf|wmf|svg)$/i.test(filename)) return;
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
