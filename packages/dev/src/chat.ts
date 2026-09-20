import { spawn, type ChildProcess } from 'node:child_process';
import { dirname, resolve } from 'node:path';
import type { IncomingMessage, ServerResponse } from 'node:http';

export type Provider = 'claude' | 'codex';
export interface FocusContext {
  slide: number | null;
  count: number;
  revision: number;
  text: string;
  entry: string;
  files: string[];
  buildError: string | null;
}
interface Message {
  role: 'user' | 'assistant';
  text: string;
  context?: FocusContext;
}

/** One shared conversation and one editing process per dev server. */
export function createChat(entry: string, notify: () => void) {
  let messages: Message[] = [];
  let provider: Provider = 'claude';
  let child: ChildProcess | undefined;
  let status = 'Ready';
  let stopping = false;
  let killTimer: ReturnType<typeof setTimeout> | undefined;
  let completion: Promise<void> = Promise.resolve();
  const snapshot = () => ({ messages, provider, running: !!child, status });
  function stop() {
    if (!child || stopping) return;
    stopping = true;
    status = 'Stopping…';
    const processToStop = child;
    const kill = (signal: NodeJS.Signals) => {
      try {
        if (process.platform !== 'win32' && processToStop.pid)
          process.kill(-processToStop.pid, signal);
        else processToStop.kill(signal);
      } catch {
        /* Already exited. */
      }
    };
    kill('SIGTERM');
    killTimer = setTimeout(() => kill('SIGKILL'), 2000);
    notify();
  }
  function start(text: string, selected: Provider, context: FocusContext) {
    provider = selected;
    const user: Message = { role: 'user', text, context };
    // Keep recent conversational intent, including each turn's original focus.
    const history = messages.slice(-12).map((message) => ({
      ...message,
      text: message.text.slice(-12000),
    }));
    const prompt = `You are editing a local PowerPoint TSX project in ${dirname(resolve(entry))}.
Read the project's authoring instructions. Make focused patches to the TSX source, preserving unrelated work.
The preview rebuilds automatically on save; do not start another dev server.
The attached focus is context, NOT a restriction: use it for references such as "this slide". For other slides or deck-wide requests, find the relevant source or shared theme instead. Do not rewrite the whole deck for a local edit.
Slide numbers are 1-based and refer to the last successful preview at send time. Source may have changed since then; verify it before editing. Files are dependency candidates, not an exact slide-to-source mapping.
Respond in the user's language, briefly describing edits or answering their question. If permissions prevent an action, explain that instead of claiming success.
Recent conversation (JSON): ${JSON.stringify(history)}
Current request and preview context (JSON): ${JSON.stringify(user)}`;
    messages.push(user);
    const answer: Message = { role: 'assistant', text: '' };
    messages.push(answer);
    if (messages.length > 100) messages = messages.slice(-100);
    status = 'Working…';
    stopping = false;
    const args =
      selected === 'codex'
        ? ['exec', '--json', '--sandbox', 'workspace-write', '--skip-git-repo-check', '-']
        : [
            '--print',
            '--verbose',
            '--output-format',
            'stream-json',
            '--permission-mode',
            'acceptEdits',
            '--tools',
            'Read,Edit,Write,Glob,Grep',
          ];
    child = spawn(selected, args, {
      cwd: dirname(resolve(entry)),
      stdio: ['pipe', 'pipe', 'pipe'],
      detached: process.platform !== 'win32',
      shell: false,
    });
    const running = child;
    let buffer = '',
      stderr = '',
      failure = '',
      sawResult = false;
    function append(value: string) {
      answer.text = (answer.text + (answer.text ? '\n\n' : '') + value).slice(-100000);
      notify();
    }
    function consume(line: string) {
      if (!line.trim()) return;
      try {
        const event = JSON.parse(line);
        if (selected === 'codex') {
          if (event.type === 'item.completed' && event.item?.type === 'agent_message')
            append(event.item.text ?? '');
          if (event.type === 'item.started') {
            status = 'Editing / inspecting…';
            notify();
          }
          if (event.type === 'turn.completed') sawResult = true;
          if (event.type === 'turn.failed' || event.type === 'error')
            failure = event.error?.message ?? event.message ?? 'Agent failed';
        } else {
          if (event.type === 'assistant') {
            for (const block of event.message?.content ?? []) {
              if (block.type === 'text') append(block.text);
              if (block.type === 'tool_use') {
                status = 'Using ' + block.name + '…';
                notify();
              }
            }
          }
          if (event.type === 'result') {
            sawResult = true;
            if (event.is_error)
              failure = event.result || event.errors?.join('\n') || 'Agent failed';
            else if (!answer.text && event.result) append(event.result);
            if (event.permission_denials?.length)
              failure = 'Some actions were denied by CLI permissions. ' + (event.result ?? '');
          }
        }
      } catch {
        failure = 'Unexpected CLI output. Update your CLI and try again.';
      }
    }
    running.stdout!.setEncoding('utf8');
    running.stdout!.on('data', (chunk: string) => {
      buffer += chunk;
      let end: number;
      while ((end = buffer.indexOf('\n')) !== -1) {
        consume(buffer.slice(0, end));
        buffer = buffer.slice(end + 1);
      }
      if (buffer.length > 2_000_000) {
        failure = 'CLI output exceeded the limit.';
        stop();
      }
    });
    running.stderr!.setEncoding('utf8');
    running.stderr!.on('data', (chunk: string) => {
      stderr = (stderr + chunk).slice(-8000);
    });
    running.stdin!.on('error', () => {
      /* Early process exit is handled below. */
    });
    running.stdin!.end(prompt);
    completion = new Promise<void>((done) => {
      running.on('error', (error) => {
        failure = `Could not run ${selected}. Install it and log in from your terminal first. ${error.message}`;
      });
      running.on('close', (code) => {
        if (buffer.trim()) consume(buffer);
        clearTimeout(killTimer);
        status = stopping
          ? 'Stopped. Saved edits remain in the preview.'
          : failure ||
            (code !== 0
              ? stderr || `${selected} exited with code ${code}`
              : !sawResult
                ? 'CLI exited without a completed result.'
                : 'Done');
        if (status !== 'Done') append(status);
        child = undefined;
        notify();
        done();
      });
    });
    notify();
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
    if (request.method === 'GET' && request.url === '/chat') {
      json(200, snapshot());
      return;
    }
    if (request.method !== 'POST') {
      json(405, { error: 'Method not allowed' });
      return;
    }
    // Same-origin JSON prevents other sites from invoking local editing tools.
    if (
      request.headers.origin !== `http://${request.headers.host}` ||
      request.headers['content-type']?.split(';')[0] !== 'application/json'
    ) {
      json(403, { error: 'Same-origin JSON required' });
      return;
    }
    try {
      let body = '';
      request.setEncoding('utf8');
      for await (const chunk of request) {
        body += chunk;
        if (Buffer.byteLength(body) > 32000) {
          json(413, { error: 'Message too large' });
          return;
        }
      }
      const value = JSON.parse(body);
      if (request.url === '/chat/stop') {
        stop();
        json(200, snapshot());
        return;
      }
      if (child) {
        json(409, { error: 'An edit is already running' });
        return;
      }
      if (request.url === '/chat/reset') {
        messages = [];
        status = 'Ready';
        notify();
        json(200, snapshot());
        return;
      }
      if (request.url !== '/chat') {
        json(404, { error: 'Not found' });
        return;
      }
      if (
        typeof value.message !== 'string' ||
        !value.message.trim() ||
        value.message.length > 16000 ||
        !['claude', 'codex'].includes(value.provider) ||
        !(value.slide === null || (Number.isInteger(value.slide) && value.slide >= 0)) ||
        !Number.isInteger(value.revision)
      ) {
        json(400, { error: 'Invalid chat request' });
        return;
      }
      const context = focus(value.slide, value.revision);
      start(value.message.trim(), value.provider, context);
      json(202, snapshot());
    } catch (cause) {
      json(400, { error: cause instanceof Error ? cause.message : 'Invalid request' });
    }
  }
  return {
    handle,
    async close() {
      stop();
      await completion;
    },
  };
}
