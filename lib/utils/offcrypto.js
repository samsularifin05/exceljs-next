'use strict';

const crypto = require('crypto');
const saxes = require('saxes');

const CFB_MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
const FREESECT = 0xffffffff;
const ENDOFCHAIN = 0xfffffffe;
const FATSECT = 0xfffffffd;
const DIFSECT = 0xfffffffc;
const NOSTREAM = 0xffffffff;

const BLOCK_KEY_VERIFIER_INPUT = Buffer.from([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]);
const BLOCK_KEY_VERIFIER_VALUE = Buffer.from([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]);
const BLOCK_KEY_ENCRYPTED_KEY = Buffer.from([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]);
const BLOCK_KEY_HMAC_KEY = Buffer.from([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]);
const BLOCK_KEY_HMAC_VALUE = Buffer.from([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]);

function isCfbBuffer(buffer) {
  return Buffer.isBuffer(buffer) && buffer.length >= 512 && buffer.subarray(0, 8).equals(CFB_MAGIC);
}

function normalizeHashAlg(alg) {
  const lower = String(alg || 'SHA512').toLowerCase().replace(/-/g, '');
  if (lower === 'sha1' || lower === 'sha256' || lower === 'sha384' || lower === 'sha512') {
    return lower;
  }
  throw new Error(`Unsupported hash algorithm: ${alg}`);
}

function cipherName(alg, bits, chaining) {
  const base = String(alg || 'AES').toLowerCase();
  if (base !== 'aes') {
    throw new Error(`Unsupported cipher algorithm: ${alg}`);
  }
  const mode = String(chaining || 'ChainingModeCBC').includes('ECB') ? 'ecb' : 'cbc';
  return `aes-${bits}-${mode}`;
}

function hashConcat(alg, ...buffers) {
  const h = crypto.createHash(alg);
  for (let i = 0; i < buffers.length; i++) {
    h.update(buffers[i]);
  }
  return h.digest();
}

function padOrTruncate(buf, size) {
  if (buf.length === size) {
    return buf;
  }
  if (buf.length > size) {
    return buf.subarray(0, size);
  }
  const out = Buffer.alloc(size, 0);
  buf.copy(out);
  return out;
}

// =============================================================================
// Minimal MS-CFB (OLE2) v3/v4 Reader & v3 Writer
// =============================================================================

function parseCfb(buffer) {
  if (!isCfbBuffer(buffer)) {
    throw new Error('Not a valid Compound File Binary (CFB) container.');
  }

  const sectorShift = buffer.readUInt16LE(30);
  const sectorSize = 1 << sectorShift;
  const miniSectorShift = buffer.readUInt16LE(32);
  const miniSectorSize = 1 << miniSectorShift;
  const numFatSectors = buffer.readUInt32LE(44);
  const firstDirSector = buffer.readUInt32LE(48);
  const miniStreamCutoff = buffer.readUInt32LE(56);
  const firstMiniFatSector = buffer.readUInt32LE(60);
  const numMiniFatSectors = buffer.readUInt32LE(64);
  const firstDifatSector = buffer.readUInt32LE(68);
  const numDifatSectors = buffer.readUInt32LE(72);

  const sectorOffset = sid => (sid + 1) * sectorSize;
  const readSector = sid => {
    const off = sectorOffset(sid);
    return buffer.subarray(off, off + sectorSize);
  };

  // Collect FAT sector IDs from header DIFAT (109 entries) + DIFAT chain
  const fatSectors = [];
  for (let i = 0; i < 109 && fatSectors.length < numFatSectors; i++) {
    const sid = buffer.readUInt32LE(76 + i * 4);
    if (sid !== FREESECT) {
      fatSectors.push(sid);
    }
  }
  let difatSid = firstDifatSector;
  const difatEntriesPerSector = sectorSize / 4 - 1;
  for (let d = 0; d < numDifatSectors && difatSid < 0xfffffffa; d++) {
    const sec = readSector(difatSid);
    for (let i = 0; i < difatEntriesPerSector && fatSectors.length < numFatSectors; i++) {
      const sid = sec.readUInt32LE(i * 4);
      if (sid !== FREESECT) {
        fatSectors.push(sid);
      }
    }
    difatSid = sec.readUInt32LE(difatEntriesPerSector * 4);
  }

  // Build FAT table
  const fat = [];
  for (let i = 0; i < fatSectors.length; i++) {
    const sec = readSector(fatSectors[i]);
    for (let j = 0; j < sectorSize; j += 4) {
      fat.push(sec.readUInt32LE(j));
    }
  }

  const readChain = (startSid, maxLen) => {
    const chunks = [];
    const visited = new Set();
    let sid = startSid;
    while (sid !== ENDOFCHAIN && sid !== FREESECT && sid < fat.length) {
      if (visited.has(sid)) {
        throw new Error('Corrupt CFB FAT chain cycle detected.');
      }
      visited.add(sid);
      chunks.push(readSector(sid));
      sid = fat[sid];
    }
    const combined = Buffer.concat(chunks);
    return maxLen !== undefined ? combined.subarray(0, maxLen) : combined;
  };

  // Read Directory entries
  const dirBuf = readChain(firstDirSector);
  const entries = [];
  for (let off = 0; off + 128 <= dirBuf.length; off += 128) {
    const nameLen = dirBuf.readUInt16LE(off + 64);
    const type = dirBuf.readUInt8(off + 66);
    if (type === 0 || nameLen < 2) {
      continue;
    }
    const name = dirBuf.subarray(off, off + Math.min(nameLen - 2, 62)).toString('utf16le');
    const startSector = dirBuf.readUInt32LE(off + 116);
    const sizeLow = dirBuf.readUInt32LE(off + 120);
    const sizeHigh = sectorSize === 4096 ? dirBuf.readUInt32LE(off + 124) : 0;
    const size = sizeHigh * 0x100000000 + sizeLow;
    entries.push({name, type, startSector, size});
  }

  const root = entries.find(e => e.type === 5);
  const miniStream = root && root.size > 0 ? readChain(root.startSector, root.size) : Buffer.alloc(0);

  // Build MiniFAT
  const miniFat = [];
  if (numMiniFatSectors > 0 && firstMiniFatSector < 0xfffffffa) {
    const miniFatBuf = readChain(firstMiniFatSector);
    for (let j = 0; j < miniFatBuf.length; j += 4) {
      miniFat.push(miniFatBuf.readUInt32LE(j));
    }
  }

  const readMiniChain = (startSid, size) => {
    const chunks = [];
    const visited = new Set();
    let sid = startSid;
    while (sid !== ENDOFCHAIN && sid !== FREESECT && sid < miniFat.length) {
      if (visited.has(sid)) {
        throw new Error('Corrupt CFB MiniFAT chain cycle detected.');
      }
      visited.add(sid);
      const off = sid * miniSectorSize;
      chunks.push(miniStream.subarray(off, off + miniSectorSize));
      sid = miniFat[sid];
    }
    return Buffer.concat(chunks).subarray(0, size);
  };

  const streams = Object.create(null);
  entries.forEach(entry => {
    if (entry.type === 2) {
      const data =
        entry.size < miniStreamCutoff
          ? readMiniChain(entry.startSector, entry.size)
          : readChain(entry.startSector, entry.size);
      streams[entry.name.toLowerCase()] = data;
    }
  });

  return streams;
}

function buildCfb(streamList) {
  const sectorSize = 512;
  const miniSectorSize = 64;
  const miniCutoff = 4096;

  // Separate mini-streams (< 4096) and regular streams (>= 4096)
  const miniChunks = [];
  const miniFat = [];
  let miniSectorCursor = 0;

  const entryMeta = streamList.map(item => {
    const data = item.data;
    if (data.length < miniCutoff) {
      const startMini = miniSectorCursor;
      const numMini = Math.max(1, Math.ceil(data.length / miniSectorSize));
      const padded = Buffer.alloc(numMini * miniSectorSize, 0);
      data.copy(padded);
      miniChunks.push(padded);
      for (let i = 0; i < numMini; i++) {
        miniFat.push(i === numMini - 1 ? ENDOFCHAIN : miniSectorCursor + i + 1);
      }
      miniSectorCursor += numMini;
      return {name: item.name, size: data.length, isMini: true, startSector: startMini};
    }
    return {name: item.name, size: data.length, isMini: false, data};
  });

  const miniStreamBuf = Buffer.concat(miniChunks);
  const miniFatSectorsCount = miniFat.length > 0 ? Math.ceil((miniFat.length * 4) / sectorSize) : 0;
  const miniStreamSectorsCount =
    miniStreamBuf.length > 0 ? Math.ceil(miniStreamBuf.length / sectorSize) : 0;

  // Allocate sectors sequentially:
  // [Directory Sectors] [MiniFAT Sectors] [MiniStream Sectors] [Regular Stream Sectors] [FAT Sectors] [DIFAT Sectors]
  const fat = [];
  const sectors = [];

  const allocChain = buf => {
    if (buf.length === 0) {
      return ENDOFCHAIN;
    }
    const count = Math.ceil(buf.length / sectorSize);
    const startSid = sectors.length;
    for (let i = 0; i < count; i++) {
      const sec = Buffer.alloc(sectorSize, 0);
      buf.copy(sec, 0, i * sectorSize, Math.min((i + 1) * sectorSize, buf.length));
      sectors.push(sec);
      fat.push(i === count - 1 ? ENDOFCHAIN : startSid + i + 1);
    }
    return startSid;
  };

  // Reserve 1 Directory sector (supports up to 4 entries: Root + up to 3 streams)
  const firstDirSector = sectors.length;
  sectors.push(Buffer.alloc(sectorSize, 0));
  fat.push(ENDOFCHAIN);

  // MiniFAT sectors
  let firstMiniFatSector = ENDOFCHAIN;
  if (miniFatSectorsCount > 0) {
    const miniFatBuf = Buffer.alloc(miniFatSectorsCount * sectorSize, 0xff);
    for (let i = 0; i < miniFat.length; i++) {
      miniFatBuf.writeUInt32LE(miniFat[i] >>> 0, i * 4);
    }
    firstMiniFatSector = allocChain(miniFatBuf);
  }

  // MiniStream sectors
  const rootStartSector = miniStreamSectorsCount > 0 ? allocChain(miniStreamBuf) : ENDOFCHAIN;

  // Regular stream sectors
  entryMeta.forEach(meta => {
    if (!meta.isMini) {
      meta.startSector = allocChain(meta.data);
    }
  });

  // Compute FAT + DIFAT sector counts iteratively
  const entriesPerFat = sectorSize / 4; // 128
  const difatEntriesPerSector = entriesPerFat - 1; // 127
  let numFatSectors = Math.ceil(fat.length / entriesPerFat);
  let numDifatSectors = 0;
  // eslint-disable-next-line no-constant-condition
  while (true) {
    const neededDifat = numFatSectors > 109 ? Math.ceil((numFatSectors - 109) / difatEntriesPerSector) : 0;
    const totalEntries = fat.length + numFatSectors + neededDifat;
    const neededFat = Math.ceil(totalEntries / entriesPerFat);
    if (neededFat === numFatSectors && neededDifat === numDifatSectors) {
      break;
    }
    numFatSectors = neededFat;
    numDifatSectors = neededDifat;
  }

  const fatSectorIds = [];
  for (let i = 0; i < numFatSectors; i++) {
    fatSectorIds.push(sectors.length);
    sectors.push(Buffer.alloc(sectorSize, 0xff));
    fat.push(FATSECT);
  }

  const difatSectorIds = [];
  for (let i = 0; i < numDifatSectors; i++) {
    difatSectorIds.push(sectors.length);
    sectors.push(Buffer.alloc(sectorSize, 0xff));
    fat.push(DIFSECT);
  }

  // Write FAT contents
  for (let i = 0; i < numFatSectors; i++) {
    const sec = sectors[fatSectorIds[i]];
    for (let j = 0; j < entriesPerFat; j++) {
      const idx = i * entriesPerFat + j;
      sec.writeUInt32LE((idx < fat.length ? fat[idx] : FREESECT) >>> 0, j * 4);
    }
  }

  // Write DIFAT sector chain if needed
  for (let i = 0; i < numDifatSectors; i++) {
    const sec = sectors[difatSectorIds[i]];
    for (let j = 0; j < difatEntriesPerSector; j++) {
      const fatIdx = 109 + i * difatEntriesPerSector + j;
      sec.writeUInt32LE((fatIdx < fatSectorIds.length ? fatSectorIds[fatIdx] : FREESECT) >>> 0, j * 4);
    }
    const nextDifat = i + 1 < numDifatSectors ? difatSectorIds[i + 1] : ENDOFCHAIN;
    sec.writeUInt32LE(nextDifat >>> 0, difatEntriesPerSector * 4);
  }

  // Write Directory entries into sectors[firstDirSector]
  const writeDirEntry = (buf, slot, opts) => {
    const off = slot * 128;
    const nameBuf = Buffer.from(`${opts.name}\0`, 'utf16le');
    nameBuf.copy(buf, off);
    buf.writeUInt16LE(nameBuf.length, off + 64);
    buf.writeUInt8(opts.type, off + 66);
    buf.writeUInt8(opts.color !== undefined ? opts.color : 1, off + 67); // 1 = black
    buf.writeUInt32LE((opts.leftSibling !== undefined ? opts.leftSibling : NOSTREAM) >>> 0, off + 68);
    buf.writeUInt32LE((opts.rightSibling !== undefined ? opts.rightSibling : NOSTREAM) >>> 0, off + 72);
    buf.writeUInt32LE((opts.child !== undefined ? opts.child : NOSTREAM) >>> 0, off + 76);
    buf.writeUInt32LE(opts.startSector >>> 0, off + 116);
    buf.writeUInt32LE(opts.size >>> 0, off + 120);
  };

  const dirSec = sectors[firstDirSector];
  // Sort entries by CFB ordering (name length first, then uppercase UTF-16LE)
  const sorted = entryMeta.slice().sort((a, b) => {
    if (a.name.length !== b.name.length) {
      return a.name.length - b.name.length;
    }
    return a.name.toUpperCase().localeCompare(b.name.toUpperCase());
  });

  writeDirEntry(dirSec, 0, {
    name: 'Root Entry',
    type: 5,
    child: sorted.length > 0 ? 1 : NOSTREAM,
    startSector: rootStartSector,
    size: miniStreamBuf.length,
  });

  sorted.forEach((meta, idx) => {
    const slot = idx + 1;
    writeDirEntry(dirSec, slot, {
      name: meta.name,
      type: 2,
      rightSibling: idx + 1 < sorted.length ? slot + 1 : NOSTREAM,
      startSector: meta.startSector,
      size: meta.size,
    });
  });

  // Build 512-byte CFB v3 Header
  const header = Buffer.alloc(sectorSize, 0);
  CFB_MAGIC.copy(header, 0);
  header.writeUInt16LE(3, 24); // minorVersion
  header.writeUInt16LE(3, 26); // majorVersion = 3
  header.writeUInt16LE(0xfffe, 28); // byteOrder little-endian
  header.writeUInt16LE(9, 30); // sectorShift = 9 (512 bytes)
  header.writeUInt16LE(6, 32); // miniSectorShift = 6 (64 bytes)
  header.writeUInt32LE(numFatSectors, 44);
  header.writeUInt32LE(firstDirSector, 48);
  header.writeUInt32LE(miniCutoff, 56);
  header.writeUInt32LE(firstMiniFatSector >>> 0, 60);
  header.writeUInt32LE(miniFatSectorsCount, 64);
  header.writeUInt32LE((numDifatSectors > 0 ? difatSectorIds[0] : ENDOFCHAIN) >>> 0, 68);
  header.writeUInt32LE(numDifatSectors, 72);
  for (let i = 0; i < 109; i++) {
    header.writeUInt32LE((i < fatSectorIds.length ? fatSectorIds[i] : FREESECT) >>> 0, 76 + i * 4);
  }

  return Buffer.concat([header, ...sectors]);
}

// =============================================================================
// MS-OFFCRYPTO Agile & Standard Encryption / Decryption
// =============================================================================

function deriveAgileHash(password, saltBuf, spinCount, hashAlg) {
  const alg = normalizeHashAlg(hashAlg);
  let h = hashConcat(alg, saltBuf, Buffer.from(password, 'utf16le'));
  const iterBuf = Buffer.alloc(4);
  for (let i = 0; i < spinCount; i++) {
    iterBuf.writeUInt32LE(i, 0);
    h = hashConcat(alg, iterBuf, h);
  }
  return h;
}

function deriveAgileKey(baseHash, blockKey, keyBytes, hashAlg) {
  const alg = normalizeHashAlg(hashAlg);
  return padOrTruncate(hashConcat(alg, baseHash, blockKey), keyBytes);
}

function aesCbcDecrypt(key, iv, data) {
  const decipher = crypto.createDecipheriv(`aes-${key.length * 8}-cbc`, key, iv);
  decipher.setAutoPadding(false);
  return Buffer.concat([decipher.update(data), decipher.final()]);
}

function aesCbcEncrypt(key, iv, data) {
  const cipher = crypto.createCipheriv(`aes-${key.length * 8}-cbc`, key, iv);
  cipher.setAutoPadding(false);
  return Buffer.concat([cipher.update(data), cipher.final()]);
}

function parseEncryptionXml(xmlText) {
  const parser = new saxes.SaxesParser();
  const info = {keyData: null, dataIntegrity: null, passwordKey: null};
  parser.on('opentag', tag => {
    if (tag.name === 'keyData') {
      info.keyData = tag.attributes;
    } else if (tag.name === 'dataIntegrity') {
      info.dataIntegrity = tag.attributes;
    } else if (tag.name === 'p:encryptedKey' || tag.name === 'encryptedKey') {
      info.passwordKey = tag.attributes;
    }
  });
  parser.write(xmlText).close();
  return info;
}

function decryptAgile(encryptionInfoBuf, encryptedPackageBuf, password) {
  const xmlText = encryptionInfoBuf.subarray(8).toString('utf8');
  const {keyData, passwordKey} = parseEncryptionXml(xmlText);
  if (!keyData || !passwordKey) {
    throw new Error('Invalid Agile EncryptionInfo XML.');
  }

  const pHashAlg = normalizeHashAlg(passwordKey.hashAlgorithm);
  const pSalt = Buffer.from(passwordKey.saltValue, 'base64');
  const pSpinCount = parseInt(passwordKey.spinCount, 10);
  const pKeyBytes = parseInt(passwordKey.keyBits, 10) / 8;
  const pBlockSize = parseInt(passwordKey.blockSize, 10) || 16;

  const baseHash = deriveAgileHash(password, pSalt, pSpinCount, pHashAlg);
  const ivPass = padOrTruncate(pSalt, pBlockSize);

  const verifierInputKey = deriveAgileKey(baseHash, BLOCK_KEY_VERIFIER_INPUT, pKeyBytes, pHashAlg);
  const verifierInputPlain = aesCbcDecrypt(
    verifierInputKey,
    ivPass,
    Buffer.from(passwordKey.encryptedVerifierHashInput, 'base64')
  ).subarray(0, pSalt.length);

  const verifierValueKey = deriveAgileKey(baseHash, BLOCK_KEY_VERIFIER_VALUE, pKeyBytes, pHashAlg);
  const expectedVerifierHash = aesCbcDecrypt(
    verifierValueKey,
    ivPass,
    Buffer.from(passwordKey.encryptedVerifierHashValue, 'base64')
  );

  const actualVerifierHash = hashConcat(pHashAlg, verifierInputPlain);
  if (!actualVerifierHash.equals(expectedVerifierHash.subarray(0, actualVerifierHash.length))) {
    throw new Error('Incorrect password for encrypted workbook.');
  }

  const secretKeyKey = deriveAgileKey(baseHash, BLOCK_KEY_ENCRYPTED_KEY, pKeyBytes, pHashAlg);
  const keyDataBytes = parseInt(keyData.keyBits, 10) / 8;
  const secretKey = aesCbcDecrypt(
    secretKeyKey,
    ivPass,
    Buffer.from(passwordKey.encryptedKeyValue, 'base64')
  ).subarray(0, keyDataBytes);

  const kdHashAlg = normalizeHashAlg(keyData.hashAlgorithm);
  const kdSalt = Buffer.from(keyData.saltValue, 'base64');
  const kdBlockSize = parseInt(keyData.blockSize, 10) || 16;

  const totalSizeLow = encryptedPackageBuf.readUInt32LE(0);
  const totalSizeHigh = encryptedPackageBuf.readUInt32LE(4);
  const totalSize = totalSizeHigh * 0x100000000 + totalSizeLow;

  const payload = encryptedPackageBuf.subarray(8);
  const segmentSize = 4096;
  const numSegments = Math.ceil(payload.length / segmentSize);
  const decryptedChunks = [];
  const segIndexBuf = Buffer.alloc(4);

  for (let i = 0; i < numSegments; i++) {
    segIndexBuf.writeUInt32LE(i, 0);
    const segIv = padOrTruncate(hashConcat(kdHashAlg, kdSalt, segIndexBuf), kdBlockSize);
    const segData = payload.subarray(i * segmentSize, (i + 1) * segmentSize);
    decryptedChunks.push(aesCbcDecrypt(secretKey, segIv, segData));
  }

  return Buffer.concat(decryptedChunks).subarray(0, totalSize);
}

function deriveStandardKey(password, saltBuf, keyBits) {
  let h = hashConcat('sha1', saltBuf, Buffer.from(password, 'utf16le'));
  const iterBuf = Buffer.alloc(4);
  for (let i = 0; i < 50000; i++) {
    iterBuf.writeUInt32LE(i, 0);
    h = hashConcat('sha1', iterBuf, h);
  }
  const blockBuf = Buffer.alloc(4, 0);
  h = hashConcat('sha1', h, blockBuf);

  const buf1 = Buffer.alloc(64, 0x36);
  for (let i = 0; i < h.length; i++) {
    buf1[i] ^= h[i];
  }
  const x1 = hashConcat('sha1', buf1);

  const buf2 = Buffer.alloc(64, 0x5c);
  for (let i = 0; i < h.length; i++) {
    buf2[i] ^= h[i];
  }
  const x2 = hashConcat('sha1', buf2);

  return Buffer.concat([x1, x2]).subarray(0, keyBits / 8);
}

function decryptStandard(encryptionInfoBuf, encryptedPackageBuf, password) {
  const headerSize = encryptionInfoBuf.readUInt32LE(8);
  const keyBits = encryptionInfoBuf.readUInt32LE(12 + 12); // EncryptionHeader.keySize
  const verifierOffset = 12 + headerSize;
  const saltSize = encryptionInfoBuf.readUInt32LE(verifierOffset);
  const salt = encryptionInfoBuf.subarray(verifierOffset + 4, verifierOffset + 4 + saltSize);
  const encryptedVerifier = encryptionInfoBuf.subarray(verifierOffset + 20, verifierOffset + 36);
  const verifierHashSize = encryptionInfoBuf.readUInt32LE(verifierOffset + 36);
  const encryptedVerifierHash = encryptionInfoBuf.subarray(verifierOffset + 40, verifierOffset + 72);

  const key = deriveStandardKey(password, salt, keyBits);
  const dec1 = crypto.createDecipheriv(cipherName('AES', keyBits, 'ECB'), key, null);
  dec1.setAutoPadding(false);
  const verifier = Buffer.concat([dec1.update(encryptedVerifier), dec1.final()]);

  const dec2 = crypto.createDecipheriv(cipherName('AES', keyBits, 'ECB'), key, null);
  dec2.setAutoPadding(false);
  const verifierHash = Buffer.concat([dec2.update(encryptedVerifierHash), dec2.final()]);

  const computedHash = hashConcat('sha1', verifier).subarray(0, verifierHashSize);
  if (!computedHash.equals(verifierHash.subarray(0, verifierHashSize))) {
    throw new Error('Incorrect password for encrypted workbook.');
  }

  const totalSizeLow = encryptedPackageBuf.readUInt32LE(0);
  const totalSizeHigh = encryptedPackageBuf.readUInt32LE(4);
  const totalSize = totalSizeHigh * 0x100000000 + totalSizeLow;

  const pkgDec = crypto.createDecipheriv(cipherName('AES', keyBits, 'ECB'), key, null);
  pkgDec.setAutoPadding(false);
  const decrypted = Buffer.concat([pkgDec.update(encryptedPackageBuf.subarray(8)), pkgDec.final()]);
  return decrypted.subarray(0, totalSize);
}

function decryptWorkbookBuffer(cfbBuffer, password) {
  if (!password) {
    throw new Error('Workbook is password-encrypted. Provide options.password to open it.');
  }
  const streams = parseCfb(cfbBuffer);
  const encryptionInfoBuf = streams.encryptioninfo;
  const encryptedPackageBuf = streams.encryptedpackage;
  if (!encryptionInfoBuf || !encryptedPackageBuf) {
    throw new Error('Missing EncryptionInfo or EncryptedPackage stream in CFB container.');
  }

  const versionMajor = encryptionInfoBuf.readUInt16LE(0);
  const versionMinor = encryptionInfoBuf.readUInt16LE(2);
  if (versionMajor === 4 && versionMinor === 4) {
    return decryptAgile(encryptionInfoBuf, encryptedPackageBuf, password);
  }
  if ((versionMajor === 3 || versionMajor === 4) && versionMinor === 2) {
    return decryptStandard(encryptionInfoBuf, encryptedPackageBuf, password);
  }
  throw new Error(`Unsupported MS-OFFCRYPTO version ${versionMajor}.${versionMinor}.`);
}

function encryptWorkbookBuffer(zipBuffer, password) {
  if (!password) {
    return zipBuffer;
  }

  const spinCount = 100000;
  const keyDataSalt = crypto.randomBytes(16);
  const keyEncryptorSalt = crypto.randomBytes(16);
  const secretKey = crypto.randomBytes(32); // AES-256
  const verifierInput = crypto.randomBytes(16);
  const verifierHash = hashConcat('sha512', verifierInput);

  const baseHash = deriveAgileHash(password, keyEncryptorSalt, spinCount, 'sha512');
  const verifierInputKey = deriveAgileKey(baseHash, BLOCK_KEY_VERIFIER_INPUT, 32, 'sha512');
  const verifierValueKey = deriveAgileKey(baseHash, BLOCK_KEY_VERIFIER_VALUE, 32, 'sha512');
  const encryptedKeyKey = deriveAgileKey(baseHash, BLOCK_KEY_ENCRYPTED_KEY, 32, 'sha512');

  const encVerifierInput = aesCbcEncrypt(verifierInputKey, keyEncryptorSalt, verifierInput);
  const encVerifierValue = aesCbcEncrypt(verifierValueKey, keyEncryptorSalt, verifierHash);
  const encKeyValue = aesCbcEncrypt(encryptedKeyKey, keyEncryptorSalt, secretKey);

  // Encrypt package in 4096-byte segments
  const sizeHeader = Buffer.alloc(8);
  sizeHeader.writeUInt32LE(zipBuffer.length >>> 0, 0);
  sizeHeader.writeUInt32LE(Math.floor(zipBuffer.length / 0x100000000), 4);

  const segmentSize = 4096;
  const numSegments = Math.max(1, Math.ceil(zipBuffer.length / segmentSize));
  const encryptedSegments = [sizeHeader];
  const segIndexBuf = Buffer.alloc(4);

  for (let i = 0; i < numSegments; i++) {
    segIndexBuf.writeUInt32LE(i, 0);
    const segIv = padOrTruncate(hashConcat('sha512', keyDataSalt, segIndexBuf), 16);
    const chunk = zipBuffer.subarray(i * segmentSize, (i + 1) * segmentSize);
    const paddedLen = Math.ceil(chunk.length / 16) * 16;
    const paddedChunk = chunk.length === paddedLen ? chunk : padOrTruncate(chunk, paddedLen);
    encryptedSegments.push(aesCbcEncrypt(secretKey, segIv, paddedChunk));
  }

  const encryptedPackageBuf = Buffer.concat(encryptedSegments);

  // DataIntegrity HMAC-SHA512
  const hmacKey = crypto.randomBytes(64);
  const hmacValue = crypto.createHmac('sha512', hmacKey).update(encryptedPackageBuf).digest();
  const hmacKeyIv = padOrTruncate(hashConcat('sha512', keyDataSalt, BLOCK_KEY_HMAC_KEY), 16);
  const hmacValIv = padOrTruncate(hashConcat('sha512', keyDataSalt, BLOCK_KEY_HMAC_VALUE), 16);
  const encryptedHmacKey = aesCbcEncrypt(secretKey, hmacKeyIv, hmacKey);
  const encryptedHmacValue = aesCbcEncrypt(secretKey, hmacValIv, hmacValue);

  const xml =
    `<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" ` +
    `xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" ` +
    `xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">` +
    `<keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" ` +
    `cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" ` +
    `saltValue="${keyDataSalt.toString('base64')}"/>` +
    `<dataIntegrity encryptedHmacKey="${encryptedHmacKey.toString('base64')}" ` +
    `encryptedHmacValue="${encryptedHmacValue.toString('base64')}"/>` +
    `<keyEncryptors>` +
    `<keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">` +
    `<p:encryptedKey spinCount="${spinCount}" saltSize="16" blockSize="16" keyBits="256" hashSize="64" ` +
    `cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" ` +
    `saltValue="${keyEncryptorSalt.toString('base64')}" ` +
    `encryptedVerifierHashInput="${encVerifierInput.toString('base64')}" ` +
    `encryptedVerifierHashValue="${encVerifierValue.toString('base64')}" ` +
    `encryptedKeyValue="${encKeyValue.toString('base64')}"/>` +
    `</keyEncryptor>` +
    `</keyEncryptors>` +
    `</encryption>`;

  const infoHeader = Buffer.alloc(8);
  infoHeader.writeUInt16LE(4, 0); // versionMajor = 4
  infoHeader.writeUInt16LE(4, 2); // versionMinor = 4
  infoHeader.writeUInt32LE(0x00000040, 4); // fAgile
  const encryptionInfoBuf = Buffer.concat([infoHeader, Buffer.from(xml, 'utf8')]);

  return buildCfb([
    {name: 'EncryptionInfo', data: encryptionInfoBuf},
    {name: 'EncryptedPackage', data: encryptedPackageBuf},
  ]);
}

module.exports = {
  isCfbBuffer,
  encryptWorkbookBuffer,
  decryptWorkbookBuffer,
};
