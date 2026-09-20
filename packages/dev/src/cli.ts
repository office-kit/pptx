#!/usr/bin/env node
import { resolve } from 'node:path';
import { exportDeck } from './index.ts';
import { serveDeck } from './server.ts';
import { initProject } from './init.ts';
import { inspectTemplate } from './inspect.ts';

const usage = `Usage: office-pptx init <new-directory>
       office-pptx dev <deck.tsx> [--port 4173]
       office-pptx build <deck.tsx> [--out deck.pptx]
       office-pptx inspect <template.pptx>`;
const [command, entry, ...args] = process.argv.slice(2);
try {
  if (command === '--help' || command === '-h') {
    console.log(usage);
  } else {
    if (!entry) throw new Error(usage);
    const option = command === 'build' ? '--out' : command === 'dev' ? '--port' : undefined;
    if (args.length && (args.length !== 2 || !option || args[0] !== option || !args[1])) {
      throw new Error(`Unexpected arguments: ${args.join(' ')}\n${usage}`);
    }
    if (command === 'init') {
      const directory = await initProject(entry);
      console.log(`Created ${directory}\nRun npm install, then npm run dev inside that directory.`);
    } else if (command === 'inspect') {
      console.log(JSON.stringify(await inspectTemplate(entry), null, 2));
    } else if (command === 'build') {
      const output = resolve(args[1] ?? 'deck.pptx');
      const result = await exportDeck(entry, output);
      console.log(`Wrote ${output} (${result.slides.length} slides)`);
    } else if (command === 'dev') {
      const port = args[1] === undefined ? 4173 : Number(args[1]);
      if (!Number.isInteger(port) || port < 0 || port > 65535)
        throw new Error('Port must be an integer between 0 and 65535.');
      const server = await serveDeck(entry, port);
      console.log(`Preview: ${server.url}`);
      const stop = () => {
        void server.close().then(() => process.exit(0));
      };
      process.once('SIGINT', stop);
      process.once('SIGTERM', stop);
    } else {
      throw new Error(usage);
    }
  }
} catch (cause) {
  console.error(cause instanceof Error ? cause.message : cause);
  process.exitCode = 1;
}
