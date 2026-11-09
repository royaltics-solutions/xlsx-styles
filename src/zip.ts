/**
 * Minimal ZIP file creator for XLSX files
 * Based on ZIP specification: https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT
 */

interface FileEntry {
  path: string;
  data: Uint8Array;
  lastModified?: Date;
}

export class MiniZip {
  private files: FileEntry[] = [];
  
  addFile(path: string, content: string | Uint8Array): void {
    const data = typeof content === 'string' 
      ? new TextEncoder().encode(content)
      : content;
    
    this.files.push({
      path,
      data,
      lastModified: new Date()
    });
  }
  
  private crc32(data: Uint8Array): number {
    let crc = 0xffffffff;
    const table = this.getCRC32Table();
    
    for (let i = 0; i < data.length; i++) {
      crc = (crc >>> 8) ^ table[(crc ^ data[i]) & 0xff];
    }
    
    return (crc ^ 0xffffffff) >>> 0;
  }
  
  private getCRC32Table(): Uint32Array {
    const table = new Uint32Array(256);
    
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let j = 0; j < 8; j++) {
        c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
      }
      table[i] = c >>> 0;
    }
    
    return table;
  }
  
  private dosDateTime(date: Date): { date: number; time: number } {
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const hours = date.getHours();
    const minutes = date.getMinutes();
    const seconds = Math.floor(date.getSeconds() / 2);
    
    const dosDate = ((year - 1980) << 9) | (month << 5) | day;
    const dosTime = (hours << 11) | (minutes << 5) | seconds;
    
    return { date: dosDate, time: dosTime };
  }
  
  generate(): Uint8Array {
    const centralDirectory: Uint8Array[] = [];
    const localFiles: Uint8Array[] = [];
    let offset = 0;
    
    // Create local file headers and file data
    for (const file of this.files) {
      const fileName = new TextEncoder().encode(file.path);
      const { date, time } = this.dosDateTime(file.lastModified || new Date());
      const crc = this.crc32(file.data);
      
      // Local file header
      const header = new Uint8Array(30 + fileName.length);
      const view = new DataView(header.buffer);
      
      // Local file header signature
      view.setUint32(0, 0x04034b50, true);
      // Version needed to extract
      view.setUint16(4, 20, true);
      // General purpose bit flag
      view.setUint16(6, 0, true);
      // Compression method (0 = no compression)
      view.setUint16(8, 0, true);
      // Last mod file time
      view.setUint16(10, time, true);
      // Last mod file date
      view.setUint16(12, date, true);
      // CRC-32
      view.setUint32(14, crc, true);
      // Compressed size
      view.setUint32(18, file.data.length, true);
      // Uncompressed size
      view.setUint32(22, file.data.length, true);
      // File name length
      view.setUint16(26, fileName.length, true);
      // Extra field length
      view.setUint16(28, 0, true);
      // File name
      header.set(fileName, 30);
      
      localFiles.push(header);
      localFiles.push(file.data);
      
      // Central directory header
      const cdHeader = new Uint8Array(46 + fileName.length);
      const cdView = new DataView(cdHeader.buffer);
      
      // Central file header signature
      cdView.setUint32(0, 0x02014b50, true);
      // Version made by
      cdView.setUint16(4, 20, true);
      // Version needed to extract
      cdView.setUint16(6, 20, true);
      // General purpose bit flag
      cdView.setUint16(8, 0, true);
      // Compression method
      cdView.setUint16(10, 0, true);
      // Last mod file time
      cdView.setUint16(12, time, true);
      // Last mod file date
      cdView.setUint16(14, date, true);
      // CRC-32
      cdView.setUint32(16, crc, true);
      // Compressed size
      cdView.setUint32(20, file.data.length, true);
      // Uncompressed size
      cdView.setUint32(24, file.data.length, true);
      // File name length
      cdView.setUint16(28, fileName.length, true);
      // Extra field length
      cdView.setUint16(30, 0, true);
      // File comment length
      cdView.setUint16(32, 0, true);
      // Disk number start
      cdView.setUint16(34, 0, true);
      // Internal file attributes
      cdView.setUint16(36, 0, true);
      // External file attributes
      cdView.setUint32(38, 0, true);
      // Relative offset of local header
      cdView.setUint32(42, offset, true);
      // File name
      cdHeader.set(fileName, 46);
      
      centralDirectory.push(cdHeader);
      offset += header.length + file.data.length;
    }
    
    // End of central directory record
    const centralDirSize = centralDirectory.reduce((sum, arr) => sum + arr.length, 0);
    const eocd = new Uint8Array(22);
    const eocdView = new DataView(eocd.buffer);
    
    // End of central dir signature
    eocdView.setUint32(0, 0x06054b50, true);
    // Number of this disk
    eocdView.setUint16(4, 0, true);
    // Number of the disk with the start of the central directory
    eocdView.setUint16(6, 0, true);
    // Total number of entries in the central directory on this disk
    eocdView.setUint16(8, this.files.length, true);
    // Total number of entries in the central directory
    eocdView.setUint16(10, this.files.length, true);
    // Size of the central directory
    eocdView.setUint32(12, centralDirSize, true);
    // Offset of start of central directory
    eocdView.setUint32(16, offset, true);
    // ZIP file comment length
    eocdView.setUint16(20, 0, true);
    
    // Combine all parts
    const totalSize = localFiles.reduce((sum, arr) => sum + arr.length, 0) + centralDirSize + 22;
    const zipFile = new Uint8Array(totalSize);
    let currentOffset = 0;
    
    // Copy local files
    for (const part of localFiles) {
      zipFile.set(part, currentOffset);
      currentOffset += part.length;
    }
    
    // Copy central directory
    for (const part of centralDirectory) {
      zipFile.set(part, currentOffset);
      currentOffset += part.length;
    }
    
    // Copy end of central directory
    zipFile.set(eocd, currentOffset);
    
    return zipFile;
  }
}

/** Tope por entrada descomprimida. Evita que un ZIP pequeño expanda cientos de MB. */
export const MAX_UNCOMPRESSED_ENTRY_BYTES = 64 * 1024 * 1024;
/** Tope acumulado de todo lo que se decodifica en una instancia. */
export const MAX_UNCOMPRESSED_TOTAL_BYTES = 256 * 1024 * 1024;

export class UnzipLimitError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'UnzipLimitError';
  }
}

interface LocalEntry {
  compressionMethod: number;
  generalPurpose: number;
  compressedSize: number;
  fileName: string;
  dataStart: number;
}

type InflateRawSync = (
  data: Uint8Array,
  options: { maxOutputLength: number },
) => Uint8Array;

function isOutputLimitError(err: unknown): boolean {
  return (
    typeof err === 'object' &&
    err !== null &&
    'code' in err &&
    (err as { code?: string }).code === 'ERR_BUFFER_TOO_LARGE'
  );
}

export interface UnzipLimits {
  /** Tope por entrada. Default: MAX_UNCOMPRESSED_ENTRY_BYTES. */
  maxEntryBytes?: number;
  /** Tope acumulado de la instancia. Default: MAX_UNCOMPRESSED_TOTAL_BYTES. */
  maxTotalBytes?: number;
}

export class MiniUnzip {
  private data: Uint8Array;
  private view: DataView;
  private readonly maxEntryBytes: number;
  private readonly maxTotalBytes: number;
  private uncompressedRemaining: number;
  
  constructor(data: Uint8Array, limits?: UnzipLimits) {
    const ab = new ArrayBuffer(data.byteLength);
    new Uint8Array(ab).set(data);
    this.data = new Uint8Array(ab);
    this.view = new DataView(ab);
    this.maxEntryBytes = limits?.maxEntryBytes ?? MAX_UNCOMPRESSED_ENTRY_BYTES;
    this.maxTotalBytes = limits?.maxTotalBytes ?? MAX_UNCOMPRESSED_TOTAL_BYTES;
    this.uncompressedRemaining = this.maxTotalBytes;
  }

  private takeUncompressed(byteLength: number): void {
    if (byteLength > this.maxEntryBytes || byteLength > this.uncompressedRemaining) {
      throw new UnzipLimitError(
        `El archivo descomprimido supera el límite permitido (${this.maxEntryBytes} bytes por entrada, ${this.maxTotalBytes} en total)`,
      );
    }
    this.uncompressedRemaining -= byteLength;
  }

  /**
   * Lee un local file header solo si el tamaño declarado cabe en el buffer.
   * Un tamaño inflado o un header truncado no avanza la lectura fuera del archivo.
   */
  private readLocalEntry(offset: number): LocalEntry | null {
    if (offset < 0 || offset + 30 > this.data.length) return null;

    const compressionMethod = this.view.getUint16(offset + 8, true);
    const generalPurpose = this.view.getUint16(offset + 6, true);
    const compressedSize = this.view.getUint32(offset + 18, true);
    const fileNameLength = this.view.getUint16(offset + 26, true);
    const extraFieldLength = this.view.getUint16(offset + 28, true);
    const nameStart = offset + 30;
    const dataStart = nameStart + fileNameLength + extraFieldLength;

    if (dataStart > this.data.length) return null;
    if (compressedSize > this.data.length - dataStart) return null;

    const fileName = new TextDecoder().decode(this.data.subarray(nameStart, nameStart + fileNameLength));
    return { compressionMethod, generalPurpose, compressedSize, fileName, dataStart };
  }
  
  static async inflateRaw(
    data: Uint8Array,
    maxOutputLength: number = MAX_UNCOMPRESSED_ENTRY_BYTES,
  ): Promise<Uint8Array> {
    if (maxOutputLength < 1) {
      throw new UnzipLimitError('Límite de descompresión inválido');
    }

    let inflateRawSync: InflateRawSync | null = null;
    try {
      const zlib = await import('zlib');
      inflateRawSync = zlib.inflateRawSync as InflateRawSync;
    } catch {
      inflateRawSync = null;
    }

    if (inflateRawSync) {
      try {
        return inflateRawSync(data, { maxOutputLength });
      } catch (err) {
        if (isOutputLimitError(err)) {
          throw new UnzipLimitError(
            `La entrada comprimida supera el límite de ${maxOutputLength} bytes`,
          );
        }
        throw err;
      }
    }

    const ds = new DecompressionStream('deflate-raw');
    const writer = ds.writable.getWriter();
    const ab = new ArrayBuffer(data.byteLength);
    new Uint8Array(ab).set(data);
    await writer.write(ab);
    await writer.close();
    const chunks: Uint8Array[] = [];
    const reader = ds.readable.getReader();
    let total = 0;
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.byteLength;
      if (total > maxOutputLength) {
        await reader.cancel();
        throw new UnzipLimitError(
          `La entrada comprimida supera el límite de ${maxOutputLength} bytes`,
        );
      }
      chunks.push(value);
    }
    const result = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks) {
      result.set(chunk, offset);
      offset += chunk.byteLength;
    }
    return result;
  }
  
  getFile(path: string): string | null {
    let offset = 0;
    
    while (offset < this.data.length - 4) {
      const signature = this.view.getUint32(offset, true);
      
      if (signature === 0x04034b50) {
        const entry = this.readLocalEntry(offset);
        if (!entry) break;

        if (entry.fileName === path) {
          if (entry.compressionMethod !== 0) return null;
          const fileData = this.data.subarray(entry.dataStart, entry.dataStart + entry.compressedSize);
          this.takeUncompressed(fileData.byteLength);
          return new TextDecoder().decode(fileData);
        }
        
        offset = entry.dataStart + entry.compressedSize;
      } else if (signature === 0x02014b50 || signature === 0x06054b50) {
        break;
      } else {
        offset++;
      }
    }
    
    return null;
  }
  
  async getFileAsync(path: string): Promise<string | null> {
    let offset = 0;
    
    while (offset < this.data.length - 4) {
      const signature = this.view.getUint32(offset, true);
      
      if (signature === 0x04034b50) {
        const entry = this.readLocalEntry(offset);
        if (!entry) break;

        const matches = entry.fileName === path;
        let consumed = entry.compressedSize;

        if (matches && entry.compressionMethod === 0) {
          const fileData = this.data.subarray(entry.dataStart, entry.dataStart + entry.compressedSize);
          this.takeUncompressed(fileData.byteLength);
          return new TextDecoder().decode(fileData);
        }

        if (matches && entry.compressionMethod === 8) {
          const compressedData = this.data.subarray(entry.dataStart, entry.dataStart + entry.compressedSize);
          const cap = Math.min(this.maxEntryBytes, this.uncompressedRemaining);
          const decompressed = await MiniUnzip.inflateRaw(compressedData, cap);
          this.takeUncompressed(decompressed.byteLength);
          return new TextDecoder().decode(decompressed);
        }

        offset = entry.dataStart + consumed;

        // Skip data descriptor if present (bit 3 set in general purpose flag)
        if (entry.generalPurpose & 0x08) {
          if (offset < this.data.length - 4) {
            const nextSig = this.view.getUint32(offset, true);
            if (nextSig === 0x08074b50) {
              offset += 16;
            }
          }
        }
      } else if (signature === 0x02014b50 || signature === 0x06054b50) {
        break;
      } else {
        offset++;
      }
    }
    
    return null;
  }
  
  listFiles(): string[] {
    const files: string[] = [];
    let offset = 0;
    
    while (offset < this.data.length - 4) {
      const signature = this.view.getUint32(offset, true);
      
      if (signature === 0x04034b50) {
        const entry = this.readLocalEntry(offset);
        if (!entry) break;
        files.push(entry.fileName);
        offset = entry.dataStart + entry.compressedSize;
      } else if (signature === 0x02014b50 || signature === 0x06054b50) {
        break;
      } else {
        offset++;
      }
    }
    
    return files;
  }
}