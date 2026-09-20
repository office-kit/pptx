import { createServer, type ServerResponse } from 'node:http';
import { watch } from 'node:fs';
import { dirname, resolve, sep } from 'node:path';
import { buildDeck, type BuildResult } from './index.ts';
import { page } from './page.ts';

export async function serveDeck(entry: string, port = 4173) {
  let latest: BuildResult | undefined;
  let error: string | null = null;
  let building = false;
  let pending = false;
  let closed = false;
  const clients = new Set<ServerResponse>();
  const server = createServer((request, response) => {
    const host = request.headers.host;
    if (host !== `127.0.0.1:${actualPort}` && host !== `localhost:${actualPort}`) {
      response.writeHead(403).end();
      return;
    }
    response.setHeader('Cache-Control', 'no-store');
    if (request.url === '/events') {
      response.writeHead(200, { 'Content-Type': 'text/event-stream' });
      response.write('data: ready\n\n');
      clients.add(response);
      request.on('close', () => clients.delete(response));
    } else if (request.url === '/state') {
      response.writeHead(200, { 'Content-Type': 'application/json' });
      response.end(
        JSON.stringify({
          slides: latest?.slides ?? [],
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
  async function rebuild() {
    if (closed) return;
    if (building) {
      pending = true;
      return;
    }
    building = true;
    try {
      latest = await buildDeck(entry);
      error = null;
    } catch (cause) {
      error = cause instanceof Error ? (cause.stack ?? cause.message) : String(cause);
    } finally {
      building = false;
      for (const client of clients) client.write('data: updated\n\n');
      if (pending) {
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
    clearTimeout(timer);
    timer = setTimeout(() => {
      void rebuild();
    }, 100);
  });
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
      for (const client of clients) client.end();
      await new Promise<void>((done, reject) =>
        server.close((cause) => (cause ? reject(cause) : done())),
      );
    },
  };
}
