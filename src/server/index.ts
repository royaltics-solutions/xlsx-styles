import { Xldx } from "../index";
import type { ReadOptions } from "../types";



Xldx.prototype.toBuffer = async function(): Promise<Buffer> {
  const uint8Array = await this.toUint8Array();
  return Buffer.from(uint8Array);
};

Xldx.prototype.write = async function(filePath: string): Promise<void> {
  const buffer = await this.toBuffer();
  const fs = await import('fs/promises');
  await fs.writeFile(filePath, buffer);
};

Xldx.prototype.download = async function(filename: string = 'download.xlsx'): Promise<void> {
  await this.write(filename);
};

export async function readFile(filePath: string, options?: ReadOptions): Promise<any> {
  const fs = await import('fs/promises');
  const data = await fs.readFile(filePath);
  return Xldx.read(data, options);
}

export * from "../index";