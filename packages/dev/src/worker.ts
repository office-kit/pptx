import { parentPort, workerData } from 'node:worker_threads';
import { buildDeck } from './build.ts';

if (!parentPort) throw new Error('The build worker must run in a worker thread.');
const { entry } = workerData as { entry: string };
parentPort.postMessage(await buildDeck(entry));
