import { inflateRawSync } from 'node:zlib';

export interface ZipEntry {
  filename: string;
  compressedSize: number;
  uncompressedSize: number;
  compressionMethod: number;
  crc32: number;
  offset: number;
}

export class ZipReader {
  private entries: ZipEntry[] = [];
  private buf: Buffer;

  constructor(data: Buffer | Uint8Array) {
    this.buf = Buffer.isBuffer(data) ? data : Buffer.from(data);
    this.readCentralDirectory();
  }

  private readCentralDirectory(): void {
    // Find End of Central Directory record (scan backwards)
    let eocdOffset = -1;
    for (let i = this.buf.length - 22; i >= 0; i--) {
      if (
        this.buf[i] === 0x50 &&
        this.buf[i + 1] === 0x4b &&
        this.buf[i + 2] === 0x05 &&
        this.buf[i + 3] === 0x06
      ) {
        eocdOffset = i;
        break;
      }
    }
    if (eocdOffset === -1) {
      throw new Error('Invalid ZIP: End of Central Directory not found');
    }

    const entryCount = this.buf.readUInt16LE(eocdOffset + 10);
    let cdOffset = this.buf.readUInt32LE(eocdOffset + 16);

    for (let i = 0; i < entryCount; i++) {
      // Central directory file header signature = 0x02014b50
      const sig = this.buf.readUInt32LE(cdOffset);
      if (sig !== 0x02014b50) {
        throw new Error(`Invalid central directory entry at offset ${cdOffset}`);
      }

      const compressionMethod = this.buf.readUInt16LE(cdOffset + 10);
      const crc32 = this.buf.readUInt32LE(cdOffset + 16);
      const compressedSize = this.buf.readUInt32LE(cdOffset + 20);
      const uncompressedSize = this.buf.readUInt32LE(cdOffset + 24);
      const filenameLen = this.buf.readUInt16LE(cdOffset + 28);
      const extraLen = this.buf.readUInt16LE(cdOffset + 30);
      const commentLen = this.buf.readUInt16LE(cdOffset + 32);
      const localHeaderOffset = this.buf.readUInt32LE(cdOffset + 42);

      const rawFilename = this.buf.toString('utf8', cdOffset + 46, cdOffset + 46 + filenameLen);
      const filename = rawFilename.replace(/\\/g, '/');

      ZipReader.validateFilename(filename);

      this.entries.push({
        filename,
        compressedSize,
        uncompressedSize,
        compressionMethod,
        crc32,
        offset: localHeaderOffset,
      });

      cdOffset += 46 + filenameLen + extraLen + commentLen;
    }
  }

  private static validateFilename(filename: string): void {
    if (filename.startsWith('/')) {
      throw new Error(`Unsafe ZIP entry: absolute path "${filename}"`);
    }
    const segments = filename.split('/');
    for (const seg of segments) {
      if (seg === '..') {
        throw new Error(`Unsafe ZIP entry: path traversal "${filename}"`);
      }
    }
  }

  getEntries(): ZipEntry[] {
    return [...this.entries];
  }

  extract(filename: string): Buffer {
    const entry = this.entries.find((e) => e.filename === filename);
    if (!entry) {
      throw new Error(`Entry not found: ${filename}`);
    }
    return this.extractEntry(entry);
  }

  private extractEntry(entry: ZipEntry): Buffer {
    const offset = entry.offset;

    // Verify local file header signature
    const sig = this.buf.readUInt32LE(offset);
    if (sig !== 0x04034b50) {
      throw new Error(`Invalid local file header at offset ${offset}`);
    }

    const filenameLen = this.buf.readUInt16LE(offset + 26);
    const extraLen = this.buf.readUInt16LE(offset + 28);
    const dataOffset = offset + 30 + filenameLen + extraLen;

    const compressedData = this.buf.subarray(dataOffset, dataOffset + entry.compressedSize);

    if (entry.compressionMethod === 0) {
      // Stored (no compression)
      return Buffer.from(compressedData);
    } else if (entry.compressionMethod === 8) {
      // Deflate
      return inflateRawSync(compressedData);
    } else {
      throw new Error(`Unsupported compression method: ${entry.compressionMethod}`);
    }
  }
}
