import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { DocxDocument } from '../dist/document.js';
import { decryptPackage, encryptPackage, isEncryptedPackage } from '../dist/encryption.js';
import { isCompoundFile, readCompoundFile, writeCompoundFile } from '../dist/cfb.js';
import { sha512SingleBlock } from '../dist/sha512.js';
import { sha1SingleBlock } from '../dist/sha1.js';

const REFERENCE = new Uint8Array(readFileSync(new URL('./fixtures/agile-msoffcrypto.docx', import.meta.url)));
const REFERENCE_PASSWORD = '参考 Reference-1';
const STANDARD = new Uint8Array(readFileSync(new URL('./fixtures/standard-office2007.docx', import.meta.url)));

test('a file encrypted by another implementation opens with its password', async () => {
  assert.equal(isEncryptedPackage(REFERENCE), true);
  const doc = await DocxDocument.load(REFERENCE, { password: REFERENCE_PASSWORD });
  assert.match(doc.getParagraphs()[0].text, /^Agile 参考文件：x+$/);
  assert.equal(doc.getPackageKind().encrypted, true);
});

test('missing, wrong and tampered inputs fail with a reason the host can act on', async () => {
  await assert.rejects(DocxDocument.load(REFERENCE), (error) => error.name === 'DocumentPasswordError' && error.reason === 'required');
  await assert.rejects(DocxDocument.load(REFERENCE, { password: 'wrong' }), (error) => error.reason === 'incorrect');
  // 改 EncryptedPackage 流里的一个字节（不是改复合文件的填充区）：HMAC 校验必须拦住。
  const streams = readCompoundFile(REFERENCE);
  const payload = Uint8Array.from(streams.get('EncryptedPackage'));
  payload[payload.length - 20] ^= 0x01;
  streams.set('EncryptedPackage', payload);
  await assert.rejects(DocxDocument.load(writeCompoundFile(streams), { password: REFERENCE_PASSWORD }),
    (error) => error.reason === 'tampered');
});

test('documents saved with a password re-open, and are a well-formed Agile package', async () => {
  const doc = DocxDocument.create();
  doc.setParagraphText(0, '机密内容 confidential');
  const encrypted = await doc.toUint8Array({ password: 'P@ss 密码' });
  assert.equal(isCompoundFile(encrypted), true);
  const streams = readCompoundFile(encrypted);
  assert.deepEqual([...streams.keys()].sort(), [
    '\x06DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace', '\x06DataSpaces/DataSpaceMap',
    '\x06DataSpaces/TransformInfo/StrongEncryptionTransform/\x06Primary', '\x06DataSpaces/Version',
    'EncryptedPackage', 'EncryptionInfo',
  ]);
  assert.deepEqual([...streams.get('EncryptionInfo').subarray(0, 8)], [4, 0, 4, 0, 0x40, 0, 0, 0], 'Agile 4.4');
  assert.match(new TextDecoder().decode(streams.get('EncryptionInfo').subarray(8)), /spinCount="100000"[^>]*hashAlgorithm="SHA512"/);
  // 明文里的字不能出现在加密包里。
  assert.equal(Buffer.from(encrypted).includes(Buffer.from('confidential')), false);
  const reopened = await DocxDocument.load(encrypted, { password: 'P@ss 密码' });
  assert.equal(reopened.getParagraphs()[0].text, '机密内容 confidential');
  await assert.rejects(DocxDocument.load(encrypted, { password: 'P@ss' }), (error) => error.reason === 'incorrect');
  // 不传密码就是普通 .docx：库不会替宿主记住密码。
  assert.equal(isCompoundFile(await reopened.toUint8Array()), false);
});

test('packages larger than one 4096-byte segment encrypt and decrypt across segment boundaries', async () => {
  const zip = Uint8Array.from({ length: 4096 * 2 + 37 }, (_, index) => (index * 31) % 251);
  const encrypted = await encryptPackage(zip, 'pw', { spinCount: 10 });
  assert.deepEqual(await decryptPackage(encrypted, 'pw'), zip);
});

test('an Office 2007 Standard-encrypted file opens', async () => {
  const doc = await DocxDocument.load(STANDARD, { password: 'Password1234_' });
  assert.equal(doc.getParagraphs()[0].text, 'lorem ipsum');
  assert.equal(doc.getPackageKind().encrypted, true);
  await assert.rejects(DocxDocument.load(STANDARD, { password: 'Password1234' }), (error) => error.reason === 'incorrect');
  await assert.rejects(DocxDocument.load(STANDARD), (error) => error.reason === 'required');
  // 存盘一律是 Agile：不把文件降级到更弱的 Standard。
  const saved = readCompoundFile(await doc.toUint8Array({ password: 'pw' }));
  assert.deepEqual([...saved.get('EncryptionInfo').subarray(0, 4)], [4, 0, 4, 0]);
});

test('a Standard header with AlgID 0 is AES-128, as its flags say', async () => {
  // 2.3.2：AlgID 为 0 时算法由标志决定，fCryptoAPI + fAES 就是 AES-128。样本本身是 AES-128。
  const streams = readCompoundFile(STANDARD);
  const info = Uint8Array.from(streams.get('EncryptionInfo'));
  const view = new DataView(info.buffer);
  assert.equal(view.getUint32(20, true), 0x660e);
  view.setUint32(20, 0, true);
  streams.set('EncryptionInfo', info);
  const plain = await decryptPackage(writeCompoundFile(streams), 'Password1234_');
  assert.deepEqual([...plain.subarray(0, 2)], [0x50, 0x4b]);
});

test('unsupported encryption schemes are reported, not misread', async () => {
  // Standard 版本号，但标志里没有 fAES：那是 RC4（CryptoAPI）。
  const streams = readCompoundFile(STANDARD);
  const info = Uint8Array.from(streams.get('EncryptionInfo'));
  info[4] = 0x04;
  streams.set('EncryptionInfo', info);
  await assert.rejects(decryptPackage(writeCompoundFile(streams), 'x'), (error) => error.reason === 'unsupported');
  // 不认识的版本号。
  const other = readCompoundFile(REFERENCE);
  const agile = Uint8Array.from(other.get('EncryptionInfo'));
  agile[0] = 2; agile[2] = 9;
  other.set('EncryptionInfo', agile);
  await assert.rejects(decryptPackage(writeCompoundFile(other), 'x'), (error) => error.reason === 'unsupported');
});

test('the single-block SHA-1 used for Standard key derivation matches WebCrypto byte for byte', async () => {
  for (const length of [0, 1, 4, 20, 24, 55]) {
    const input = Uint8Array.from({ length }, (_, index) => (index * 13 + length) & 0xff);
    assert.deepEqual(sha1SingleBlock(input), new Uint8Array(await crypto.subtle.digest('SHA-1', input)), `${length} bytes`);
  }
  assert.throws(() => sha1SingleBlock(new Uint8Array(56)), /55/);
});

test('the single-block SHA-512 used for key derivation matches WebCrypto byte for byte', async () => {
  for (const length of [0, 1, 4, 55, 56, 64, 68, 100, 111]) {
    const input = Uint8Array.from({ length }, (_, index) => (index * 7 + length) & 0xff);
    const expected = new Uint8Array(await crypto.subtle.digest('SHA-512', input));
    assert.deepEqual(sha512SingleBlock(input), expected, `${length} bytes`);
  }
  assert.throws(() => sha512SingleBlock(new Uint8Array(112)), /111/);
});

test('the compound file reader rejects hostile structures instead of looping', () => {
  const valid = writeCompoundFile(new Map([['A', new Uint8Array(5000).fill(1)], ['B', new Uint8Array([1, 2, 3])]]));
  assert.deepEqual([...readCompoundFile(valid).get('B')], [1, 2, 3]);
  // FAT 里让第一个扇区指向自己：扇区链成环。
  const looped = Uint8Array.from(valid);
  const view = new DataView(looped.buffer);
  const fatSector = view.getUint32(0x4c, true);
  const fatOffset = 512 + fatSector * 512;
  view.setUint32(fatOffset + 4, 1, true);
  assert.throws(() => readCompoundFile(looped), /chain|invalid|range/i);
  // 目录起始扇区越界。
  const outOfRange = Uint8Array.from(valid);
  new DataView(outOfRange.buffer).setUint32(0x30, 99999, true);
  assert.throws(() => readCompoundFile(outOfRange), /range|invalid/i);
  assert.throws(() => readCompoundFile(new Uint8Array(600)), /Not a compound file/);
});
