// 测试用的 worker：宿主要写的就是这几行。
import { parentPort } from 'node:worker_threads';
import { compareDocxBytes, readDocxSnapshot, searchDocxText } from '../dist/offload.js';

const handlers = { compareDocxBytes, readDocxSnapshot, searchDocxText };

parentPort.on('message', async ({ method, args }) => {
  try {
    parentPort.postMessage({ ok: true, value: await handlers[method](...args) });
  } catch (error) {
    parentPort.postMessage({ ok: false, message: String(error) });
  }
});
