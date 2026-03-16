import { describe, it, expect } from 'vitest';
import { ZipReader, ZipWriter } from '../src/zip/index.js';

describe('ZipWriter + ZipReader round-trip', () => {
  it('writes and reads a single file', () => {
    const writer = new ZipWriter();
    writer.addFile('hello.txt', 'Hello, World!');

    const zipBuf = writer.toBuffer();
    const reader = new ZipReader(zipBuf);

    const entries = reader.getEntries();
    expect(entries).toHaveLength(1);
    expect(entries[0].filename).toBe('hello.txt');

    const content = reader.extract('hello.txt').toString('utf8');
    expect(content).toBe('Hello, World!');
  });

  it('writes and reads multiple files', () => {
    const writer = new ZipWriter();
    writer.addFile('a.txt', 'content A');
    writer.addFile('dir/b.txt', 'content B');
    writer.addFile('dir/c.xml', '<root><item/></root>');

    const zipBuf = writer.toBuffer();
    const reader = new ZipReader(zipBuf);

    const entries = reader.getEntries();
    expect(entries).toHaveLength(3);

    expect(reader.extract('a.txt').toString('utf8')).toBe('content A');
    expect(reader.extract('dir/b.txt').toString('utf8')).toBe('content B');
    expect(reader.extract('dir/c.xml').toString('utf8')).toBe('<root><item/></root>');
  });

  it('handles empty files', () => {
    const writer = new ZipWriter();
    writer.addFile('empty.txt', '');

    const zipBuf = writer.toBuffer();
    const reader = new ZipReader(zipBuf);

    const entries = reader.getEntries();
    expect(entries).toHaveLength(1);
    expect(reader.extract('empty.txt').toString('utf8')).toBe('');
  });

  it('handles binary data', () => {
    const writer = new ZipWriter();
    const data = Buffer.from([0, 1, 2, 255, 254, 253, 128, 127]);
    writer.addFile('binary.bin', data);

    const zipBuf = writer.toBuffer();
    const reader = new ZipReader(zipBuf);

    const extracted = reader.extract('binary.bin');
    expect(Buffer.compare(extracted, data)).toBe(0);
  });

  it('throws on missing entry', () => {
    const writer = new ZipWriter();
    writer.addFile('a.txt', 'data');

    const reader = new ZipReader(writer.toBuffer());
    expect(() => reader.extract('missing.txt')).toThrow('Entry not found');
  });

  it('handles empty archive', () => {
    const writer = new ZipWriter();
    const zipBuf = writer.toBuffer();
    const reader = new ZipReader(zipBuf);
    expect(reader.getEntries()).toHaveLength(0);
  });

  it('handles large compressible data', () => {
    const writer = new ZipWriter();
    const data = 'A'.repeat(100000);
    writer.addFile('large.txt', data);

    const zipBuf = writer.toBuffer();
    // Compressed size should be much smaller
    expect(zipBuf.length).toBeLessThan(data.length);

    const reader = new ZipReader(zipBuf);
    expect(reader.extract('large.txt').toString('utf8')).toBe(data);
  });
});
