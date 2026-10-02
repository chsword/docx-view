/**
 * 单块 SHA-1（输入不超过 55 字节）。只给 Office 2007「Standard」加密的口令派生用：那一步对
 * 「4 字节计数 + 20 字节哈希」连续算 50000 次，每次 await WebCrypto 要好几秒。结果由测试对着
 * WebCrypto 逐字节核对。SHA-1 在这里只是规范规定的派生步骤，不是我们选的安全原语。
 */
const w = new Int32Array(80);
const block = new Uint8Array(64);

export function sha1SingleBlock(input: Uint8Array): Uint8Array {
  if (input.length > 55) throw new Error('sha1SingleBlock handles at most 55 bytes.');
  block.fill(0);
  block.set(input);
  block[input.length] = 0x80;
  const bits = input.length * 8;
  block[62] = (bits >>> 8) & 0xff;
  block[63] = bits & 0xff;
  for (let index = 0; index < 16; index++) {
    w[index] = (block[index * 4]! << 24) | (block[index * 4 + 1]! << 16) | (block[index * 4 + 2]! << 8) | block[index * 4 + 3]!;
  }
  for (let index = 16; index < 80; index++) {
    const value = w[index - 3]! ^ w[index - 8]! ^ w[index - 14]! ^ w[index - 16]!;
    w[index] = (value << 1) | (value >>> 31);
  }
  let a = 0x67452301;
  let b = 0xefcdab89;
  let c = 0x98badcfe;
  let d = 0x10325476;
  let e = 0xc3d2e1f0;
  for (let index = 0; index < 80; index++) {
    const [f, k] = index < 20 ? [(b & c) | (~b & d), 0x5a827999]
      : index < 40 ? [b ^ c ^ d, 0x6ed9eba1]
        : index < 60 ? [(b & c) | (b & d) | (c & d), 0x8f1bbcdc]
          : [b ^ c ^ d, 0xca62c1d6];
    const temp = (((a << 5) | (a >>> 27)) + f + e + k + w[index]!) | 0;
    e = d;
    d = c;
    c = (b << 30) | (b >>> 2);
    b = a;
    a = temp;
  }
  const state = [a + 0x67452301, b + 0xefcdab89, c + 0x98badcfe, d + 0x10325476, e + 0xc3d2e1f0];
  const output = new Uint8Array(20);
  state.forEach((value, index) => {
    for (let byte = 0; byte < 4; byte++) output[index * 4 + byte] = (value >>> (24 - byte * 8)) & 0xff;
  });
  return output;
}
