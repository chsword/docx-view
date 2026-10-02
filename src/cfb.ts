/**
 * 复合文件（CFB / OLE2，MS-CFB）的最小读写：加密的 .docx 不是 ZIP，而是一个复合文件，里面放
 * `EncryptionInfo` 和 `EncryptedPackage` 两个流（外加固定的 `\x06DataSpaces` 存储）。
 *
 * 读取面向**不可信输入**：所有扇区号、链长、偏移都做界限检查，扇区链用访问集合防环，目录树
 * 遍历有深度上限。写入只产出 v3（512 字节扇区），小于 4096 字节的流按规范放进迷你流。
 */

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const FREE = 0xffffffff;
const END_OF_CHAIN = 0xfffffffe;
const FAT_SECTOR = 0xfffffffd;
const DIFAT_SECTOR = 0xfffffffc;
const NO_STREAM = 0xffffffff;
const MINI_CUTOFF = 4096;
const MINI_SECTOR = 64;

export function isCompoundFile(bytes: Uint8Array): boolean {
  return bytes.length >= 512 && SIGNATURE.every((value, index) => bytes[index] === value);
}

/** 读出复合文件里的全部流，键是用 `/` 连起来的路径（如 `\x06DataSpaces/Version`）。 */
export function readCompoundFile(bytes: Uint8Array): Map<string, Uint8Array> {
  if (!isCompoundFile(bytes)) throw new Error('Not a compound file.');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = (offset: number) => view.getUint16(offset, true);
  const u32 = (offset: number) => view.getUint32(offset, true);
  const sectorShift = u16(0x1e);
  if (sectorShift !== 9 && sectorShift !== 12) throw new Error('Unsupported compound file sector size.');
  const sectorSize = 1 << sectorShift;
  const sectorCount = Math.floor((bytes.length - sectorSize) / sectorSize);
  const sectorOffset = (sector: number) => {
    if (sector >= sectorCount) throw new Error('Compound file sector out of range.');
    return sectorSize + sector * sectorSize;
  };
  const sector = (index: number) => bytes.subarray(sectorOffset(index), sectorOffset(index) + sectorSize);

  // DIFAT：头部 109 个，其余在 DIFAT 扇区链上。
  const fatSectors: number[] = [];
  for (let index = 0; index < 109; index++) {
    const value = u32(0x4c + index * 4);
    if (value !== FREE) fatSectors.push(value);
  }
  let difat = u32(0x44);
  const seenDifat = new Set<number>();
  while (difat !== END_OF_CHAIN && difat !== FREE) {
    if (seenDifat.has(difat) || seenDifat.size > sectorCount) throw new Error('Compound file DIFAT chain loops.');
    seenDifat.add(difat);
    const data = sector(difat);
    const entries = sectorSize / 4 - 1;
    const dataView = new DataView(data.buffer, data.byteOffset, data.byteLength);
    for (let index = 0; index < entries; index++) {
      const value = dataView.getUint32(index * 4, true);
      if (value !== FREE) fatSectors.push(value);
    }
    difat = dataView.getUint32(entries * 4, true);
  }
  const fat = new Uint32Array(fatSectors.length * (sectorSize / 4));
  fatSectors.forEach((fatSector, index) => {
    const data = sector(fatSector);
    fat.set(new Uint32Array(data.buffer.slice(data.byteOffset, data.byteOffset + sectorSize)), index * (sectorSize / 4));
  });
  const chain = (start: number, table: Uint32Array, limit: number): number[] => {
    const result: number[] = [];
    const seen = new Set<number>();
    for (let current = start; current !== END_OF_CHAIN; current = table[current] ?? END_OF_CHAIN) {
      if (current === FREE || current === FAT_SECTOR || current === DIFAT_SECTOR) break;
      if (current >= table.length || seen.has(current) || result.length > limit) throw new Error('Compound file sector chain is invalid.');
      seen.add(current);
      result.push(current);
    }
    return result;
  };
  const readChain = (start: number, size?: number): Uint8Array => {
    const sectors = chain(start, fat, sectorCount);
    const output = new Uint8Array(sectors.length * sectorSize);
    sectors.forEach((index, position) => output.set(sector(index), position * sectorSize));
    return size === undefined ? output : output.subarray(0, Math.min(size, output.length));
  };

  const directory = readChain(u32(0x30));
  const entryCount = Math.floor(directory.length / 128);
  const directoryView = new DataView(directory.buffer, directory.byteOffset, directory.byteLength);
  const entry = (index: number) => {
    const base = index * 128;
    const nameLength = Math.min(64, directoryView.getUint16(base + 0x40, true));
    let name = '';
    for (let offset = 0; offset + 2 < nameLength; offset += 2) name += String.fromCharCode(directoryView.getUint16(base + offset, true));
    return {
      name,
      type: directory[base + 0x42]!,
      left: directoryView.getUint32(base + 0x44, true),
      right: directoryView.getUint32(base + 0x48, true),
      child: directoryView.getUint32(base + 0x4c, true),
      start: directoryView.getUint32(base + 0x74, true),
      size: directoryView.getUint32(base + 0x78, true),
    };
  };
  if (entryCount < 1) throw new Error('Compound file has no root entry.');
  const root = entry(0);
  const miniStream = root.start === END_OF_CHAIN ? new Uint8Array() : readChain(root.start, root.size);
  const miniFatBytes = u32(0x3c) === END_OF_CHAIN ? new Uint8Array() : readChain(u32(0x3c));
  const miniFat = new Uint32Array(miniFatBytes.buffer.slice(miniFatBytes.byteOffset, miniFatBytes.byteOffset + (miniFatBytes.byteLength & ~3)));
  const readMini = (start: number, size: number): Uint8Array => {
    const sectors = chain(start, miniFat, Math.ceil(miniStream.length / MINI_SECTOR));
    const output = new Uint8Array(sectors.length * MINI_SECTOR);
    sectors.forEach((index, position) => {
      const offset = index * MINI_SECTOR;
      if (offset + MINI_SECTOR > miniStream.length) throw new Error('Compound file mini sector out of range.');
      output.set(miniStream.subarray(offset, offset + MINI_SECTOR), position * MINI_SECTOR);
    });
    return output.subarray(0, Math.min(size, output.length));
  };

  const streams = new Map<string, Uint8Array>();
  const visited = new Set<number>();
  const walk = (index: number, prefix: string, depth: number) => {
    if (index === NO_STREAM) return;
    if (index >= entryCount || visited.has(index) || depth > 64) throw new Error('Compound file directory tree is invalid.');
    visited.add(index);
    const item = entry(index);
    walk(item.left, prefix, depth + 1);
    walk(item.right, prefix, depth + 1);
    const path = prefix ? `${prefix}/${item.name}` : item.name;
    if (item.type === 2) streams.set(path, item.size < MINI_CUTOFF ? readMini(item.start, item.size) : readChain(item.start, item.size));
    else if (item.type === 1) walk(item.child, path, depth + 1);
  };
  walk(root.child, '', 0);
  return streams;
}

/** 写一个 v3 复合文件。`streams` 的键用 `/` 表示存储层级。 */
export function writeCompoundFile(streams: Map<string, Uint8Array>): Uint8Array {
  interface Node { name: string; type: 1 | 2 | 5; children: Node[]; data?: Uint8Array; start?: number; size?: number; id?: number }
  const root: Node = { name: 'Root Entry', type: 5, children: [] };
  for (const [path, data] of streams) {
    const segments = path.split('/');
    let parent = root;
    segments.forEach((segment, index) => {
      if (segment.length > 31) throw new Error('Compound file entry name too long.');
      let child = parent.children.find((candidate) => candidate.name === segment);
      if (!child) {
        child = index === segments.length - 1 ? { name: segment, type: 2, children: [], data } : { name: segment, type: 1, children: [] };
        parent.children.push(child);
      }
      parent = child;
    });
  }
  const sectorSize = 512;
  // 迷你流：小流按 64 字节一格依次放。
  const miniChunks: Uint8Array[] = [];
  const miniFat: number[] = [];
  const bigStreams: Node[] = [];
  const nodes: Node[] = [];
  const collect = (node: Node) => {
    nodes.push(node);
    for (const child of node.children) collect(child);
  };
  collect(root);
  for (const node of nodes) {
    if (node.type !== 2 || !node.data) continue;
    node.size = node.data.length;
    if (node.data.length >= MINI_CUTOFF) { bigStreams.push(node); continue; }
    if (!node.data.length) { node.start = END_OF_CHAIN; continue; }
    const count = Math.ceil(node.data.length / MINI_SECTOR);
    node.start = miniFat.length;
    for (let index = 0; index < count; index++) miniFat.push(index === count - 1 ? END_OF_CHAIN : miniFat.length + 1);
    const padded = new Uint8Array(count * MINI_SECTOR);
    padded.set(node.data);
    miniChunks.push(padded);
  }
  const miniStream = concat(miniChunks);
  const sectorsFor = (length: number) => Math.ceil(length / sectorSize);
  // 目录：每项 128 字节，红黑树退化成「每层兄弟按名字排、全挂右兄弟链」也是合法的（全黑）。
  nodes.forEach((node, index) => { node.id = index; });
  const directoryBytes = new Uint8Array(sectorsFor(nodes.length * 128) * sectorSize);
  const plan: Array<{ data: Uint8Array; assign: (start: number) => void }> = [];
  plan.push({ data: directoryBytes, assign: () => {} });
  const miniFatBytes = new Uint8Array(sectorsFor(miniFat.length * 4) * sectorSize);
  new DataView(miniFatBytes.buffer).setUint32(0, 0, true);
  for (let index = 0; index < miniFat.length; index++) new DataView(miniFatBytes.buffer).setUint32(index * 4, miniFat[index]!, true);
  for (let index = miniFat.length; index < miniFatBytes.length / 4; index++) new DataView(miniFatBytes.buffer).setUint32(index * 4, FREE, true);
  if (miniFat.length) plan.push({ data: miniFatBytes, assign: () => {} });
  if (miniStream.length) plan.push({ data: miniStream, assign: (start) => { root.start = start; root.size = miniStream.length; } });
  for (const node of bigStreams) plan.push({ data: node.data!, assign: (start) => { node.start = start; } });
  // 先算数据扇区，再算 FAT 需要几个扇区（FAT 自己也占 FAT 项）。
  const dataSectors = plan.reduce((sum, item) => sum + sectorsFor(item.data.length), 0);
  let fatSectorCount = 1;
  while (fatSectorCount * (sectorSize / 4) < dataSectors + fatSectorCount) fatSectorCount++;
  if (fatSectorCount > 109) throw new Error('Compound file too large for a header-only DIFAT.');
  const fat = new Uint32Array(fatSectorCount * (sectorSize / 4)).fill(FREE);
  let cursor = 0;
  const starts: number[] = [];
  for (const item of plan) {
    const count = sectorsFor(item.data.length);
    starts.push(count ? cursor : END_OF_CHAIN);
    for (let index = 0; index < count; index++) fat[cursor + index] = index === count - 1 ? END_OF_CHAIN : cursor + index + 1;
    item.assign(count ? cursor : END_OF_CHAIN);
    cursor += count;
  }
  const fatStart = cursor;
  for (let index = 0; index < fatSectorCount; index++) fat[fatStart + index] = FAT_SECTOR;
  if (!miniStream.length) { root.start = END_OF_CHAIN; root.size = 0; }

  // 目录项。同一存储下的兄弟按 MS-CFB 的比较规则（先比长度，再不分大小写比名字）排好，建成一棵
  // 平衡二叉树：中间的当根，左右递归。叶子深度最多差一层，把最深那一层（不满时）涂红、其余涂黑，
  // 各路径的黑节点数就相同，是合法的红黑树——规范要求兄弟构成红黑树，只连成一条右链的写法
  // 宽松的读取方认，Word 不一定认。
  const directoryView = new DataView(directoryBytes.buffer);
  const links = new Map<number, { left: number; right: number; red: boolean }>();
  const compare = (a: Node, b: Node) => a.name.length - b.name.length || a.name.toUpperCase().localeCompare(b.name.toUpperCase());
  const balanced = (sorted: Node[]): number => {
    if (!sorted.length) return NO_STREAM;
    const height = Math.floor(Math.log2(sorted.length));
    const perfect = sorted.length === 2 ** (height + 1) - 1;
    const build = (from: number, to: number, depth: number): number => {
      if (from > to) return NO_STREAM;
      const middle = (from + to) >> 1;
      const id = sorted[middle]!.id!;
      links.set(id, { left: build(from, middle - 1, depth + 1), right: build(middle + 1, to, depth + 1), red: !perfect && depth === height });
      return id;
    };
    return build(0, sorted.length - 1, 0);
  };
  const childRoot = new Map<number, number>();
  for (const node of nodes) childRoot.set(node.id!, balanced([...node.children].sort(compare)));
  for (const node of nodes) {
    const base = node.id! * 128;
    const link = links.get(node.id!) ?? { left: NO_STREAM, right: NO_STREAM, red: false };
    for (let index = 0; index < node.name.length; index++) directoryView.setUint16(base + index * 2, node.name.charCodeAt(index), true);
    directoryView.setUint16(base + 0x40, (node.name.length + 1) * 2, true);
    directoryBytes[base + 0x42] = node.type;
    directoryBytes[base + 0x43] = link.red ? 0 : 1;
    directoryView.setUint32(base + 0x44, link.left, true);
    directoryView.setUint32(base + 0x48, link.right, true);
    directoryView.setUint32(base + 0x4c, node.type === 2 ? NO_STREAM : childRoot.get(node.id!)!, true);
    directoryView.setUint32(base + 0x74, node.type === 1 ? 0 : (node.start ?? END_OF_CHAIN), true);
    directoryView.setUint32(base + 0x78, node.type === 1 ? 0 : (node.size ?? 0), true);
  }
  for (let index = nodes.length; index < directoryBytes.length / 128; index++) {
    const base = index * 128;
    directoryView.setUint32(base + 0x44, NO_STREAM, true);
    directoryView.setUint32(base + 0x48, NO_STREAM, true);
    directoryView.setUint32(base + 0x4c, NO_STREAM, true);
  }

  const header = new Uint8Array(sectorSize);
  const headerView = new DataView(header.buffer);
  header.set(SIGNATURE);
  headerView.setUint16(0x18, 0x3e, true);
  headerView.setUint16(0x1a, 3, true);
  headerView.setUint16(0x1c, 0xfffe, true);
  headerView.setUint16(0x1e, 9, true);
  headerView.setUint16(0x20, 6, true);
  headerView.setUint32(0x2c, fatSectorCount, true);
  headerView.setUint32(0x30, starts[0]!, true);
  headerView.setUint32(0x38, MINI_CUTOFF, true);
  headerView.setUint32(0x3c, miniFat.length ? starts[1]! : END_OF_CHAIN, true);
  headerView.setUint32(0x40, miniFat.length ? sectorsFor(miniFatBytes.length) : 0, true);
  headerView.setUint32(0x44, END_OF_CHAIN, true);
  headerView.setUint32(0x48, 0, true);
  for (let index = 0; index < 109; index++) headerView.setUint32(0x4c + index * 4, index < fatSectorCount ? fatStart + index : FREE, true);

  const body = new Uint8Array((cursor + fatSectorCount) * sectorSize);
  let offset = 0;
  for (const item of plan) {
    body.set(item.data, offset);
    offset += sectorsFor(item.data.length) * sectorSize;
  }
  body.set(new Uint8Array(fat.buffer), offset);
  return concat([header, body]);
}

function concat(chunks: Uint8Array[]): Uint8Array {
  const result = new Uint8Array(chunks.reduce((sum, chunk) => sum + chunk.length, 0));
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.length;
  }
  return result;
}
