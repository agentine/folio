import { deflateRawSync } from 'node:zlib';

interface FileEntry {
  filename: string;
  data: Buffer;
  compressedData: Buffer;
  compressionMethod: number;
  crc32: number;
  offset: number;
}

// CRC32 lookup table
const crcTable: Uint32Array = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c;
  }
  return table;
})();

function crc32(data: Buffer): number {
  let crc = 0xffffffff;
  for (let i = 0; i < data.length; i++) {
    crc = crcTable[(crc ^ data[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

export class ZipWriter {
  private entries: FileEntry[] = [];

  addFile(filename: string, data: Buffer | Uint8Array | string): void {
    const buf = typeof data === 'string' ? Buffer.from(data, 'utf8') : Buffer.isBuffer(data) ? data : Buffer.from(data);

    let compressedData: Buffer;
    let compressionMethod: number;

    if (buf.length === 0) {
      // Store empty files
      compressedData = buf;
      compressionMethod = 0;
    } else {
      compressedData = deflateRawSync(buf);
      // Use deflate if it actually saves space, otherwise store
      if (compressedData.length < buf.length) {
        compressionMethod = 8;
      } else {
        compressedData = buf;
        compressionMethod = 0;
      }
    }

    this.entries.push({
      filename,
      data: buf,
      compressedData,
      compressionMethod,
      crc32: crc32(buf),
      offset: 0,
    });
  }

  toBuffer(): Buffer {
    const parts: Buffer[] = [];
    let offset = 0;

    // Write local file headers + data
    for (const entry of this.entries) {
      entry.offset = offset;
      const filenameBytes = Buffer.from(entry.filename, 'utf8');

      // Local file header (30 bytes + filename)
      const header = Buffer.alloc(30);
      header.writeUInt32LE(0x04034b50, 0);     // Signature
      header.writeUInt16LE(20, 4);              // Version needed (2.0)
      header.writeUInt16LE(0, 6);               // General purpose flags
      header.writeUInt16LE(entry.compressionMethod, 8);
      header.writeUInt16LE(0, 10);              // Mod time
      header.writeUInt16LE(0, 12);              // Mod date
      header.writeUInt32LE(entry.crc32, 14);
      header.writeUInt32LE(entry.compressedData.length, 18);
      header.writeUInt32LE(entry.data.length, 22);
      header.writeUInt16LE(filenameBytes.length, 26);
      header.writeUInt16LE(0, 28);              // Extra field length

      parts.push(header, filenameBytes, entry.compressedData);
      offset += 30 + filenameBytes.length + entry.compressedData.length;
    }

    const centralDirectoryOffset = offset;

    // Write central directory
    for (const entry of this.entries) {
      const filenameBytes = Buffer.from(entry.filename, 'utf8');

      const cdEntry = Buffer.alloc(46);
      cdEntry.writeUInt32LE(0x02014b50, 0);     // Signature
      cdEntry.writeUInt16LE(20, 4);             // Version made by
      cdEntry.writeUInt16LE(20, 6);             // Version needed
      cdEntry.writeUInt16LE(0, 8);              // Flags
      cdEntry.writeUInt16LE(entry.compressionMethod, 10);
      cdEntry.writeUInt16LE(0, 12);             // Mod time
      cdEntry.writeUInt16LE(0, 14);             // Mod date
      cdEntry.writeUInt32LE(entry.crc32, 16);
      cdEntry.writeUInt32LE(entry.compressedData.length, 20);
      cdEntry.writeUInt32LE(entry.data.length, 24);
      cdEntry.writeUInt16LE(filenameBytes.length, 28);
      cdEntry.writeUInt16LE(0, 30);             // Extra field length
      cdEntry.writeUInt16LE(0, 32);             // Comment length
      cdEntry.writeUInt16LE(0, 34);             // Disk number start
      cdEntry.writeUInt16LE(0, 36);             // Internal file attributes
      cdEntry.writeUInt32LE(0, 38);             // External file attributes
      cdEntry.writeUInt32LE(entry.offset, 42);  // Local header offset

      parts.push(cdEntry, filenameBytes);
      offset += 46 + filenameBytes.length;
    }

    const centralDirectorySize = offset - centralDirectoryOffset;

    // End of central directory record
    const eocd = Buffer.alloc(22);
    eocd.writeUInt32LE(0x06054b50, 0);          // Signature
    eocd.writeUInt16LE(0, 4);                   // Disk number
    eocd.writeUInt16LE(0, 6);                   // Central directory disk
    eocd.writeUInt16LE(this.entries.length, 8);  // Entries on disk
    eocd.writeUInt16LE(this.entries.length, 10); // Total entries
    eocd.writeUInt32LE(centralDirectorySize, 12);
    eocd.writeUInt32LE(centralDirectoryOffset, 16);
    eocd.writeUInt16LE(0, 20);                  // Comment length

    parts.push(eocd);

    return Buffer.concat(parts);
  }
}
