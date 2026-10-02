import { isCompoundFile, readCompoundFile, writeCompoundFile } from './cfb.js';
import { sha512SingleBlock } from './sha512.js';
import { sha1SingleBlock } from './sha1.js';

/**
 * 加密的 .docx（MS-OFFCRYPTO「Agile」加密，Word 2010 起的默认）：复合文件里一个 XML 的
 * `EncryptionInfo`，一个按 4096 字节分段 AES-CBC 加密的 `EncryptedPackage`（里面就是原本的 ZIP）。
 *
 * 只用 WebCrypto（`globalThis.crypto.subtle`），浏览器与 Node 都有，不加依赖。密码派生是
 * SHA-512 迭代 spinCount 次（Word 写 100000），加密分段、校验 HMAC 都按规范来；数据完整性
 * （`dataIntegrity`）在解密时**校验**——被改过的包拒绝打开，而不是交出一份被篡改的文档。
 *
 * 也能**打开** Office 2007 的「Standard」加密（AES-128/192/256 + SHA-1，没有完整性校验）；存盘一律用
 * Agile。RC4（CryptoAPI）不支持，遇到时报明确的错误。
 */

const BLOCK_KEY_VERIFIER_INPUT = [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79];
const BLOCK_KEY_VERIFIER_VALUE = [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e];
const BLOCK_KEY_ENCRYPTED_KEY = [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6];
const BLOCK_KEY_HMAC_KEY = [0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6];
const BLOCK_KEY_HMAC_VALUE = [0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33];
const SEGMENT = 4096;
const MAX_SPIN_COUNT = 10_000_000;

/** `\x06DataSpaces` 下的四个流是固定内容（MS-OFFCRYPTO 2.1 的结构），Word 靠它们认出加密包。 */
const DATA_SPACES: Array<[string, string]> = [
  ['\x06DataSpaces/Version', '3c0000004d006900630072006f0073006f00660074002e0043006f006e007400610069006e00650072002e004400610074006100530070006100630065007300010000000100000001000000'],
  ['\x06DataSpaces/DataSpaceMap', '08000000010000006800000001000000000000002000000045006e0063007200790070007400650064005000610063006b00610067006500320000005300740072006f006e00670045006e006300720079007000740069006f006e004400610074006100530070006100630065000000'],
  ['\x06DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace', '0800000001000000320000005300740072006f006e00670045006e006300720079007000740069006f006e005400720061006e00730066006f0072006d000000'],
  ['\x06DataSpaces/TransformInfo/StrongEncryptionTransform/\x06Primary', '58000000010000004c0000007b00460046003900410033004600300033002d0035003600450046002d0034003600310033002d0042004400440035002d003500410034003100430031004400300037003200340036007d004e0000004d006900630072006f0073006f00660074002e0043006f006e007400610069006e00650072002e0045006e006300720079007000740069006f006e005400720061006e00730066006f0072006d00000001000000010000000100000000000000000000000000000004000000'],
];

export class DocumentPasswordError extends Error {
  constructor(message: string, readonly reason: 'required' | 'incorrect' | 'tampered' | 'unsupported') {
    super(message);
    this.name = 'DocumentPasswordError';
  }
}

/** 是不是加密的 Office 包（复合文件里有 EncryptionInfo）。 */
export function isEncryptedPackage(bytes: Uint8Array): boolean {
  if (!isCompoundFile(bytes)) return false;
  try {
    return readCompoundFile(bytes).has('EncryptionInfo');
  } catch {
    return false;
  }
}

function subtle(): SubtleCrypto {
  const value = globalThis.crypto?.subtle;
  if (!value) throw new Error('Encrypted documents need WebCrypto (globalThis.crypto.subtle).');
  return value;
}

const toBuffer = (bytes: Uint8Array): ArrayBuffer => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

function concat(...chunks: Array<Uint8Array | number[]>): Uint8Array {
  const parts = chunks.map((chunk) => (chunk instanceof Uint8Array ? chunk : Uint8Array.from(chunk)));
  const result = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let offset = 0;
  for (const part of parts) { result.set(part, offset); offset += part.length; }
  return result;
}

const hashName = (algorithm: string) => {
  const map: Record<string, string> = { SHA512: 'SHA-512', SHA384: 'SHA-384', SHA256: 'SHA-256', SHA1: 'SHA-1' };
  const name = map[algorithm.toUpperCase().replace('-', '')];
  if (!name) throw new DocumentPasswordError(`Unsupported hash algorithm: ${algorithm}.`, 'unsupported');
  return name;
};

async function digest(algorithm: string, ...chunks: Array<Uint8Array | number[]>): Promise<Uint8Array> {
  return new Uint8Array(await subtle().digest(hashName(algorithm), toBuffer(concat(...chunks))));
}

const le32 = (value: number) => [value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff];

/** 长度对不上时按规范截断或用 0x36 补齐。 */
function fit(bytes: Uint8Array, length: number, padByte = 0x36): Uint8Array {
  if (bytes.length >= length) return bytes.subarray(0, length);
  const result = new Uint8Array(length).fill(padByte);
  result.set(bytes);
  return result;
}

async function aesKey(raw: Uint8Array, usage: KeyUsage): Promise<CryptoKey> {
  return subtle().importKey('raw', toBuffer(raw), { name: 'AES-CBC' }, false, [usage]);
}

/**
 * 不带填充的 AES-CBC 解密。WebCrypto 的 AES-CBC 一定按 PKCS#7 去填充，而 Agile 的分段没有
 * 填充：在密文后面补一个块，让它解出来正好是一整块 0x10 的填充——这个块就是用最后一个密文块作 IV
 * 把那块填充加密一次得到的第一个密文块。
 */
async function decryptRaw(key: Uint8Array, iv: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  if (data.length % 16) throw new DocumentPasswordError('Encrypted data is not block aligned.', 'tampered');
  if (!data.length) return new Uint8Array();
  const last = data.subarray(data.length - 16);
  const encryptKey = await aesKey(key, 'encrypt');
  const padding = new Uint8Array(await subtle().encrypt({ name: 'AES-CBC', iv: toBuffer(last) }, encryptKey, toBuffer(new Uint8Array(16).fill(16))));
  const decryptKey = await aesKey(key, 'decrypt');
  return new Uint8Array(await subtle().decrypt({ name: 'AES-CBC', iv: toBuffer(iv) }, decryptKey, toBuffer(concat(data, padding.subarray(0, 16)))));
}

/** 不带填充的 AES-CBC 加密：先补齐到块大小，再丢掉 WebCrypto 追加的那一块填充。 */
async function encryptRaw(key: Uint8Array, iv: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  const padded = data.length % 16 ? fit(data, Math.ceil(data.length / 16) * 16, 0) : data;
  const encrypted = new Uint8Array(await subtle().encrypt({ name: 'AES-CBC', iv: toBuffer(iv) }, await aesKey(key, 'encrypt'), toBuffer(padded)));
  return encrypted.subarray(0, padded.length);
}

async function passwordHash(password: string, salt: Uint8Array, spinCount: number, algorithm: string): Promise<Uint8Array> {
  const utf16 = new Uint8Array(password.length * 2);
  for (let index = 0; index < password.length; index++) {
    utf16[index * 2] = password.charCodeAt(index) & 0xff;
    utf16[index * 2 + 1] = password.charCodeAt(index) >>> 8;
  }
  let hash = await digest(algorithm, salt, utf16);
  if (hashName(algorithm) === 'SHA-512') {
    // 迭代的输入总是「4 字节计数 + 64 字节哈希」，一个块就装下：同步算比每次 await WebCrypto
    // 快一个数量级（100000 次从约 5 秒降到约 0.5 秒）。
    const input = new Uint8Array(68);
    for (let iteration = 0; iteration < spinCount; iteration++) {
      input[0] = iteration & 0xff;
      input[1] = (iteration >>> 8) & 0xff;
      input[2] = (iteration >>> 16) & 0xff;
      input[3] = (iteration >>> 24) & 0xff;
      input.set(hash, 4);
      hash = sha512SingleBlock(input);
    }
    return hash;
  }
  for (let iteration = 0; iteration < spinCount; iteration++) hash = await digest(algorithm, le32(iteration), hash);
  return hash;
}

async function derivedKey(hash: Uint8Array, blockKey: number[], keyBytes: number, algorithm: string): Promise<Uint8Array> {
  return fit(await digest(algorithm, hash, blockKey), keyBytes);
}

interface AgileParameters {
  keyData: { salt: Uint8Array; blockSize: number; keyBits: number; hashSize: number; hash: string };
  password: {
    spinCount: number; salt: Uint8Array; blockSize: number; keyBits: number; hashSize: number; hash: string;
    verifierInput: Uint8Array; verifierValue: Uint8Array; keyValue: Uint8Array;
  };
  integrity?: { hmacKey: Uint8Array; hmacValue: Uint8Array };
}

const base64 = (value: string | null): Uint8Array => {
  if (!value) throw new DocumentPasswordError('EncryptionInfo is missing a required value.', 'tampered');
  const binary = atob(value);
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
};
const toBase64 = (bytes: Uint8Array) => btoa(Array.from(bytes, (byte) => String.fromCharCode(byte)).join(''));

function parseEncryptionInfo(info: Uint8Array): AgileParameters {
  if (info.length < 8) throw new DocumentPasswordError('EncryptionInfo is truncated.', 'tampered');
  const major = info[0]! | (info[1]! << 8);
  const minor = info[2]! | (info[3]! << 8);
  if (major !== 4 || minor !== 4) {
    throw new DocumentPasswordError(`Only Agile encryption (version 4.4) is supported; this file uses version ${major}.${minor}.`, 'unsupported');
  }
  const xml = new TextDecoder().decode(info.subarray(8));
  // 只取需要的属性，用正则不建 DOM：这是不可信输入，结构简单，属性名固定。
  const element = (name: string) => new RegExp(`<(?:\\w+:)?${name}\\b([^>]*)>`).exec(xml)?.[1];
  const attribute = (attrs: string | undefined, name: string) => (attrs ? new RegExp(`\\b${name}="([^"]*)"`).exec(attrs)?.[1] ?? null : null);
  const number = (attrs: string | undefined, name: string) => {
    const value = Number(attribute(attrs, name));
    if (!Number.isSafeInteger(value) || value <= 0) throw new DocumentPasswordError(`EncryptionInfo has an invalid ${name}.`, 'tampered');
    return value;
  };
  const keyData = element('keyData');
  const encryptedKey = element('encryptedKey');
  if (!keyData || !encryptedKey) throw new DocumentPasswordError('Only password-protected Agile encryption is supported.', 'unsupported');
  for (const attrs of [keyData, encryptedKey]) {
    if (attribute(attrs, 'cipherAlgorithm') !== 'AES' || attribute(attrs, 'cipherChaining') !== 'ChainingModeCBC') {
      throw new DocumentPasswordError('Only AES-CBC Agile encryption is supported.', 'unsupported');
    }
  }
  const spinCount = number(encryptedKey, 'spinCount');
  if (spinCount > MAX_SPIN_COUNT) throw new DocumentPasswordError('EncryptionInfo spinCount is unreasonably large.', 'tampered');
  const integrity = element('dataIntegrity');
  return {
    keyData: {
      salt: base64(attribute(keyData, 'saltValue')), blockSize: number(keyData, 'blockSize'), keyBits: number(keyData, 'keyBits'),
      hashSize: number(keyData, 'hashSize'), hash: attribute(keyData, 'hashAlgorithm') ?? 'SHA512',
    },
    password: {
      spinCount, salt: base64(attribute(encryptedKey, 'saltValue')), blockSize: number(encryptedKey, 'blockSize'),
      keyBits: number(encryptedKey, 'keyBits'), hashSize: number(encryptedKey, 'hashSize'),
      hash: attribute(encryptedKey, 'hashAlgorithm') ?? 'SHA512',
      verifierInput: base64(attribute(encryptedKey, 'encryptedVerifierHashInput')),
      verifierValue: base64(attribute(encryptedKey, 'encryptedVerifierHashValue')),
      keyValue: base64(attribute(encryptedKey, 'encryptedKeyValue')),
    },
    ...(integrity ? { integrity: {
      hmacKey: base64(attribute(integrity, 'encryptedHmacKey')), hmacValue: base64(attribute(integrity, 'encryptedHmacValue')),
    } } : {}),
  };
}

async function hmac(algorithm: string, key: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  const imported = await subtle().importKey('raw', toBuffer(key), { name: 'HMAC', hash: hashName(algorithm) }, false, ['sign']);
  return new Uint8Array(await subtle().sign('HMAC', imported, toBuffer(data)));
}

const sameBytes = (a: Uint8Array, b: Uint8Array) => a.length === b.length && a.every((value, index) => value === b[index]);

/** 解开加密包，返回里面的 ZIP 字节。密码错、包被改过、加密方式不支持时抛 `DocumentPasswordError`。 */
export async function decryptPackage(bytes: Uint8Array, password: string | undefined): Promise<Uint8Array> {
  const streams = readCompoundFile(bytes);
  const info = streams.get('EncryptionInfo');
  const encrypted = streams.get('EncryptedPackage');
  if (!info || !encrypted) throw new DocumentPasswordError('Compound file is not an encrypted Office package.', 'unsupported');
  // 版本 3.2 / 4.2 是 Office 2007 的「Standard」加密；4.4 是 Agile。别的（RC4 等）由 parseEncryptionInfo 报不支持。
  if (info.length >= 4 && (info[0] === 3 || info[0] === 4) && info[1] === 0 && info[2] === 2 && info[3] === 0) {
    const standard = parseStandardEncryptionInfo(info);
    if (password === undefined) throw new DocumentPasswordError('This document is password protected.', 'required');
    return decryptStandard(standard, encrypted, password);
  }
  const parameters = parseEncryptionInfo(info);
  if (password === undefined) throw new DocumentPasswordError('This document is password protected.', 'required');
  const { password: key, keyData } = parameters;
  const hash = await passwordHash(password, key.salt, key.spinCount, key.hash);
  const keyBytes = key.keyBits / 8;
  const verifier = (await decryptRaw(await derivedKey(hash, BLOCK_KEY_VERIFIER_INPUT, keyBytes, key.hash), key.salt, key.verifierInput))
    .subarray(0, key.salt.length);
  const expected = (await decryptRaw(await derivedKey(hash, BLOCK_KEY_VERIFIER_VALUE, keyBytes, key.hash), key.salt, key.verifierValue))
    .subarray(0, key.hashSize);
  if (!sameBytes(await digest(key.hash, verifier), expected)) throw new DocumentPasswordError('Incorrect password.', 'incorrect');
  const secret = (await decryptRaw(await derivedKey(hash, BLOCK_KEY_ENCRYPTED_KEY, keyBytes, key.hash), key.salt, key.keyValue))
    .subarray(0, keyData.keyBits / 8);

  if (parameters.integrity) {
    const hmacKeyIv = fit(await digest(keyData.hash, keyData.salt, BLOCK_KEY_HMAC_KEY), keyData.blockSize);
    const hmacValueIv = fit(await digest(keyData.hash, keyData.salt, BLOCK_KEY_HMAC_VALUE), keyData.blockSize);
    const hmacKey = (await decryptRaw(secret, hmacKeyIv, parameters.integrity.hmacKey)).subarray(0, keyData.hashSize);
    const hmacValue = (await decryptRaw(secret, hmacValueIv, parameters.integrity.hmacValue)).subarray(0, keyData.hashSize);
    // 规范与 Word：HMAC 覆盖整个 EncryptedPackage 流（含开头 8 字节的长度）。
    if (!sameBytes(await hmac(keyData.hash, hmacKey, encrypted), hmacValue)) {
      throw new DocumentPasswordError('The encrypted document failed its integrity check.', 'tampered');
    }
  }

  if (encrypted.length < 8) throw new DocumentPasswordError('EncryptedPackage is truncated.', 'tampered');
  const view = new DataView(encrypted.buffer, encrypted.byteOffset, encrypted.byteLength);
  const size = view.getUint32(0, true) + view.getUint32(4, true) * 2 ** 32;
  const payload = encrypted.subarray(8);
  if (size > payload.length) throw new DocumentPasswordError('EncryptedPackage is shorter than its declared size.', 'tampered');
  const output = new Uint8Array(Math.ceil(payload.length / SEGMENT) * SEGMENT);
  for (let segment = 0, offset = 0; offset < payload.length; segment++, offset += SEGMENT) {
    const iv = fit(await digest(keyData.hash, keyData.salt, le32(segment)), keyData.blockSize);
    output.set(await decryptRaw(secret, iv, payload.subarray(offset, Math.min(offset + SEGMENT, payload.length))), offset);
  }
  return output.subarray(0, size);
}

interface StandardParameters {
  keyBytes: number;
  salt: Uint8Array;
  verifier: Uint8Array;
  verifierHash: Uint8Array;
}

/**
 * Standard 加密（MS-OFFCRYPTO 2.3.4.5）的二进制头：版本、标志、头长度、EncryptionHeader（算法、
 * 哈希、密钥位数、CSP 名），再接 EncryptionVerifier（盐、加密的校验值与其哈希）。只收 AES +
 * SHA-1——规范规定 Standard 加密就是这一种组合；标志里没有 fAES 的是 RC4（CryptoAPI），不支持。
 */
function parseStandardEncryptionInfo(info: Uint8Array): StandardParameters {
  const view = new DataView(info.buffer, info.byteOffset, info.byteLength);
  const need = (length: number) => {
    if (info.length < length) throw new DocumentPasswordError('EncryptionInfo is truncated.', 'tampered');
  };
  need(12);
  const flags = view.getUint32(4, true);
  if (!(flags & 0x04) || !(flags & 0x20)) {
    throw new DocumentPasswordError('Only AES Standard encryption is supported; RC4 CryptoAPI encryption is not.', 'unsupported');
  }
  const headerSize = view.getUint32(8, true);
  if (headerSize < 32 || headerSize > 1024) throw new DocumentPasswordError('EncryptionInfo has an invalid header size.', 'tampered');
  need(12 + headerSize + 40);
  const algorithm = view.getUint32(12 + 8, true);
  const hash = view.getUint32(12 + 12, true);
  const keyBits = view.getUint32(12 + 16, true);
  const keyBitsByAlgorithm: Record<number, number> = { 0x660e: 128, 0x660f: 192, 0x6610: 256 };
  if (!keyBitsByAlgorithm[algorithm] || (hash !== 0x8004 && hash !== 0)) {
    throw new DocumentPasswordError('Unsupported Standard encryption algorithm.', 'unsupported');
  }
  if (keyBits !== keyBitsByAlgorithm[algorithm]) throw new DocumentPasswordError('EncryptionInfo key size does not match its algorithm.', 'tampered');
  const verifierOffset = 12 + headerSize;
  const saltSize = view.getUint32(verifierOffset, true);
  if (saltSize !== 16) throw new DocumentPasswordError('EncryptionInfo has an invalid salt size.', 'tampered');
  const salt = info.subarray(verifierOffset + 4, verifierOffset + 20);
  const verifier = info.subarray(verifierOffset + 20, verifierOffset + 36);
  const hashSize = view.getUint32(verifierOffset + 36, true);
  if (hashSize !== 20) throw new DocumentPasswordError('EncryptionInfo has an invalid verifier hash size.', 'tampered');
  // AES 下加密的哈希补齐到 32 字节。
  need(verifierOffset + 40 + 32);
  return { keyBytes: keyBits / 8, salt, verifier, verifierHash: info.subarray(verifierOffset + 40, verifierOffset + 72) };
}

/**
 * Standard 的口令派生（MS-OFFCRYPTO 2.3.4.7）：SHA-1 迭代 50000 次，再与块号 0 一起哈希，然后
 * 按 0x36 / 0x5C 各填满 64 字节异或、各哈希一次，前后拼起来截到密钥长度。
 */
async function standardKey(password: string, parameters: StandardParameters): Promise<Uint8Array> {
  const utf16 = new Uint8Array(password.length * 2);
  for (let index = 0; index < password.length; index++) {
    utf16[index * 2] = password.charCodeAt(index) & 0xff;
    utf16[index * 2 + 1] = password.charCodeAt(index) >>> 8;
  }
  let hash = await digest('SHA1', parameters.salt, utf16);
  const input = new Uint8Array(24);
  for (let iteration = 0; iteration < 50_000; iteration++) {
    input.set(le32(iteration), 0);
    input.set(hash, 4);
    hash = sha1SingleBlock(input);
  }
  const final = await digest('SHA1', hash, le32(0));
  const xorPad = (byte: number) => {
    const buffer = new Uint8Array(64).fill(byte);
    final.forEach((value, index) => { buffer[index] = buffer[index]! ^ value; });
    return buffer;
  };
  return concat(await digest('SHA1', xorPad(0x36)), await digest('SHA1', xorPad(0x5c))).subarray(0, parameters.keyBytes);
}

/**
 * 不带填充的 AES-ECB 解密。WebCrypto 没有 ECB：用零 IV 做一次 CBC 解密，得到的是
 * D(Cᵢ) ⊕ Cᵢ₋₁，再把每块异或回前一个密文块就是 D(Cᵢ)。
 */
async function decryptEcb(key: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  const plain = await decryptRaw(key, new Uint8Array(16), data);
  for (let offset = 16; offset < plain.length; offset++) plain[offset] = plain[offset]! ^ data[offset - 16]!;
  return plain;
}

async function decryptStandard(parameters: StandardParameters, encrypted: Uint8Array, password: string): Promise<Uint8Array> {
  const key = await standardKey(password, parameters);
  const verifier = await decryptEcb(key, parameters.verifier);
  const expected = (await decryptEcb(key, parameters.verifierHash)).subarray(0, 20);
  if (!sameBytes(await digest('SHA1', verifier), expected)) throw new DocumentPasswordError('Incorrect password.', 'incorrect');
  if (encrypted.length < 8) throw new DocumentPasswordError('EncryptedPackage is truncated.', 'tampered');
  const view = new DataView(encrypted.buffer, encrypted.byteOffset, encrypted.byteLength);
  const size = view.getUint32(0, true) + view.getUint32(4, true) * 2 ** 32;
  const payload = encrypted.subarray(8, 8 + Math.floor((encrypted.length - 8) / 16) * 16);
  if (size > payload.length) throw new DocumentPasswordError('EncryptedPackage is shorter than its declared size.', 'tampered');
  // Standard 加密没有 HMAC：被改过的包只能在后面解 ZIP 时暴露。
  return (await decryptEcb(key, payload)).subarray(0, size);
}

/** 用密码把 ZIP 字节加密成 Agile 加密包（与 Word 2010+ 相同的参数：AES-256、SHA-512、100000 次）。 */
export async function encryptPackage(zip: Uint8Array, password: string, options: { spinCount?: number } = {}): Promise<Uint8Array> {
  const spinCount = options.spinCount ?? 100_000;
  if (!Number.isSafeInteger(spinCount) || spinCount < 1 || spinCount > MAX_SPIN_COUNT) throw new Error('spinCount must be 1–10000000.');
  const random = (length: number) => globalThis.crypto.getRandomValues(new Uint8Array(length));
  const keyDataSalt = random(16);
  const passwordSalt = random(16);
  const secret = random(32);
  const verifier = random(16);
  const hmacKey = random(64);
  const hash = await passwordHash(password, passwordSalt, spinCount, 'SHA512');

  const sizeHeader = new Uint8Array(8);
  const sizeView = new DataView(sizeHeader.buffer);
  sizeView.setUint32(0, zip.length % 2 ** 32, true);
  sizeView.setUint32(4, Math.floor(zip.length / 2 ** 32), true);
  const segments: Uint8Array[] = [sizeHeader];
  for (let segment = 0, offset = 0; offset < zip.length; segment++, offset += SEGMENT) {
    const iv = fit(await digest('SHA512', keyDataSalt, le32(segment)), 16);
    segments.push(await encryptRaw(secret, iv, zip.subarray(offset, Math.min(offset + SEGMENT, zip.length))));
  }
  const encryptedPackage = concat(...segments);

  const hmacKeyIv = fit(await digest('SHA512', keyDataSalt, BLOCK_KEY_HMAC_KEY), 16);
  const hmacValueIv = fit(await digest('SHA512', keyDataSalt, BLOCK_KEY_HMAC_VALUE), 16);
  const encryptedHmacKey = await encryptRaw(secret, hmacKeyIv, hmacKey);
  const encryptedHmacValue = await encryptRaw(secret, hmacValueIv, await hmac('SHA512', hmacKey, encryptedPackage));
  const encryptedVerifierInput = await encryptRaw(await derivedKey(hash, BLOCK_KEY_VERIFIER_INPUT, 32, 'SHA512'), passwordSalt, verifier);
  const encryptedVerifierValue = await encryptRaw(await derivedKey(hash, BLOCK_KEY_VERIFIER_VALUE, 32, 'SHA512'), passwordSalt, await digest('SHA512', verifier));
  const encryptedKeyValue = await encryptRaw(await derivedKey(hash, BLOCK_KEY_ENCRYPTED_KEY, 32, 'SHA512'), passwordSalt, secret);

  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    + '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">'
    + `<keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${toBase64(keyDataSalt)}"/>`
    + `<dataIntegrity encryptedHmacKey="${toBase64(encryptedHmacKey)}" encryptedHmacValue="${toBase64(encryptedHmacValue)}"/>`
    + '<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">'
    + `<p:encryptedKey spinCount="${spinCount}" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${toBase64(passwordSalt)}" `
    + `encryptedVerifierHashInput="${toBase64(encryptedVerifierInput)}" encryptedVerifierHashValue="${toBase64(encryptedVerifierValue)}" encryptedKeyValue="${toBase64(encryptedKeyValue)}"/>`
    + '</keyEncryptor></keyEncryptors></encryption>';
  const info = concat([0x04, 0x00, 0x04, 0x00, 0x40, 0x00, 0x00, 0x00], new TextEncoder().encode(xml));
  const hex = (value: string) => Uint8Array.from(value.match(/../g)!, (pair) => parseInt(pair, 16));
  return writeCompoundFile(new Map<string, Uint8Array>([
    ...DATA_SPACES.map(([path, value]) => [path, hex(value)] as [string, Uint8Array]),
    ['EncryptionInfo', info],
    ['EncryptedPackage', encryptedPackage],
  ]));
}
