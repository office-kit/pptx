import { once } from 'node:events';
import { parentPort, workerData } from 'node:worker_threads';
import type { PreviewCache } from './preview-cache.ts';
import { buildDeck } from './build.ts';

if (!parentPort) throw new Error('The build worker must run in a worker thread.');
const { directory } = workerData as { directory: string };
const [request] = (await once(parentPort, 'message')) as [
  { dependencies: string[]; cache?: PreviewCache },
];
parentPort.postMessage(await buildDeck(directory, request.dependencies, request.cache));
