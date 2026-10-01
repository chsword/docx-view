// 测试用的 worker：宿主要写的就是这几行。
import { parentPort } from 'node:worker_threads';
import { zipParts } from '../dist/zip.js';

parentPort.on('message', async (parts) => {
  try {
    parentPort.postMessage({ ok: true, bytes: await zipParts(parts) });
  } catch (error) {
    parentPort.postMessage({ ok: false, message: String(error) });
  }
});
