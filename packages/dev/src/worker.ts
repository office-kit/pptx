import { once } from 'node:events';
import { parentPort, workerData } from 'node:worker_threads';
import { buildDeck } from './build.ts';

if (!parentPort) throw new Error('The build worker must run in a worker thread.');
const { entry, directory } = workerData as { entry: string; directory: string };
await once(parentPort, 'message');
parentPort.postMessage(await buildDeck(entry, directory));
