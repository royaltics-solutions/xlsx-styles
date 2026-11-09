import { describe, it, expect } from "vitest";
import { deflateRawSync } from "zlib";
import { MiniZip, MiniUnzip, UnzipLimitError } from "../src/zip";

function concatBytes(parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

function localFile(name: string, method: number, payload: Uint8Array): Uint8Array {
  const nameBytes = new TextEncoder().encode(name);
  const out = new Uint8Array(30 + nameBytes.length + payload.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, 0x04034b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(8, method, true);
  view.setUint32(18, payload.length, true);
  view.setUint32(22, payload.length, true);
  view.setUint16(26, nameBytes.length, true);
  out.set(nameBytes, 30);
  out.set(payload, 30 + nameBytes.length);
  return out;
}

describe("MiniZip", () => {
  describe("addFile", () => {
    it("should add a string file", () => {
      const zip = new MiniZip();
      zip.addFile("test.txt", "Hello, World!");

      const data = zip.generate();
      expect(data).toBeInstanceOf(Uint8Array);
      expect(data.length).toBeGreaterThan(0);
    });

    it("should add a Uint8Array file", () => {
      const zip = new MiniZip();
      const content = new Uint8Array([72, 101, 108, 108, 111]); // "Hello"
      zip.addFile("binary.bin", content);

      const data = zip.generate();
      expect(data).toBeInstanceOf(Uint8Array);
    });

    it("should add multiple files", () => {
      const zip = new MiniZip();
      zip.addFile("file1.txt", "Content 1");
      zip.addFile("file2.txt", "Content 2");
      zip.addFile("file3.txt", "Content 3");

      const data = zip.generate();
      expect(data.length).toBeGreaterThan(0);
    });

    it("should handle nested paths", () => {
      const zip = new MiniZip();
      zip.addFile("folder/subfolder/file.txt", "Nested content");

      const data = zip.generate();
      expect(data.length).toBeGreaterThan(0);
    });

    it("should handle empty content", () => {
      const zip = new MiniZip();
      zip.addFile("empty.txt", "");

      const data = zip.generate();
      expect(data.length).toBeGreaterThan(0);
    });

    it("should handle unicode content", () => {
      const zip = new MiniZip();
      zip.addFile("unicode.txt", "Hello 世界 🌍");

      const data = zip.generate();
      expect(data.length).toBeGreaterThan(0);
    });
  });

  describe("generate", () => {
    it("should generate valid ZIP signature", () => {
      const zip = new MiniZip();
      zip.addFile("test.txt", "test");

      const data = zip.generate();

      // Local file header signature: PK\x03\x04
      expect(data[0]).toBe(0x50); // P
      expect(data[1]).toBe(0x4b); // K
      expect(data[2]).toBe(0x03);
      expect(data[3]).toBe(0x04);
    });

    it("should include central directory", () => {
      const zip = new MiniZip();
      zip.addFile("test.txt", "test content");

      const data = zip.generate();

      // Look for central directory signature: PK\x01\x02
      let foundCentralDir = false;
      for (let i = 0; i < data.length - 4; i++) {
        if (data[i] === 0x50 && data[i + 1] === 0x4b &&
            data[i + 2] === 0x01 && data[i + 3] === 0x02) {
          foundCentralDir = true;
          break;
        }
      }
      expect(foundCentralDir).toBe(true);
    });

    it("should include end of central directory", () => {
      const zip = new MiniZip();
      zip.addFile("test.txt", "test");

      const data = zip.generate();

      // End of central directory signature: PK\x05\x06
      const lastBytes = data.slice(-22);
      expect(lastBytes[0]).toBe(0x50); // P
      expect(lastBytes[1]).toBe(0x4b); // K
      expect(lastBytes[2]).toBe(0x05);
      expect(lastBytes[3]).toBe(0x06);
    });

    it("should handle large files", () => {
      const zip = new MiniZip();
      const largeContent = "x".repeat(100000);
      zip.addFile("large.txt", largeContent);

      const data = zip.generate();
      expect(data.length).toBeGreaterThan(100000);
    });

    it("should generate empty ZIP when no files added", () => {
      const zip = new MiniZip();
      const data = zip.generate();

      // Should still have valid ZIP structure with EOCD
      expect(data.length).toBe(22); // Just EOCD record
    });
  });
});

describe("MiniUnzip", () => {
  describe("getFile", () => {
    it("should retrieve added file content", () => {
      const zip = new MiniZip();
      zip.addFile("test.txt", "Hello, World!");

      const data = zip.generate();
      const unzip = new MiniUnzip(data);

      const content = unzip.getFile("test.txt");
      expect(content).toBe("Hello, World!");
    });

    it("should retrieve specific file from multiple", () => {
      const zip = new MiniZip();
      zip.addFile("file1.txt", "Content 1");
      zip.addFile("file2.txt", "Content 2");
      zip.addFile("file3.txt", "Content 3");

      const data = zip.generate();
      const unzip = new MiniUnzip(data);

      expect(unzip.getFile("file1.txt")).toBe("Content 1");
      expect(unzip.getFile("file2.txt")).toBe("Content 2");
      expect(unzip.getFile("file3.txt")).toBe("Content 3");
    });

    it("should return null for non-existent file", () => {
      const zip = new MiniZip();
      zip.addFile("exists.txt", "content");

      const data = zip.generate();
      const unzip = new MiniUnzip(data);

      const content = unzip.getFile("doesnotexist.txt");
      expect(content).toBeNull();
    });

    it("should handle nested paths", () => {
      const zip = new MiniZip();
      zip.addFile("folder/subfolder/file.txt", "Nested!");

      const data = zip.generate();
      const unzip = new MiniUnzip(data);

      expect(unzip.getFile("folder/subfolder/file.txt")).toBe("Nested!");
    });

    it("should handle empty content", () => {
      const zip = new MiniZip();
      zip.addFile("empty.txt", "");

      const data = zip.generate();
      const unzip = new MiniUnzip(data);

      expect(unzip.getFile("empty.txt")).toBe("");
    });

    it("should handle unicode content", () => {
      const zip = new MiniZip();
      zip.addFile("unicode.txt", "日本語 한국어 中文");

      const data = zip.generate();
      const unzip = new MiniUnzip(data);

      expect(unzip.getFile("unicode.txt")).toBe("日本語 한국어 中文");
    });
  });

  describe("listFiles", () => {
    it("should list all files in ZIP", () => {
      const zip = new MiniZip();
      zip.addFile("a.txt", "A");
      zip.addFile("b.txt", "B");
      zip.addFile("c.txt", "C");

      const data = zip.generate();
      const unzip = new MiniUnzip(data);

      const files = unzip.listFiles();
      expect(files).toContain("a.txt");
      expect(files).toContain("b.txt");
      expect(files).toContain("c.txt");
      expect(files).toHaveLength(3);
    });

    it("should list files with nested paths", () => {
      const zip = new MiniZip();
      zip.addFile("root.txt", "root");
      zip.addFile("folder/nested.txt", "nested");
      zip.addFile("folder/deep/deeper.txt", "deep");

      const data = zip.generate();
      const unzip = new MiniUnzip(data);

      const files = unzip.listFiles();
      expect(files).toContain("root.txt");
      expect(files).toContain("folder/nested.txt");
      expect(files).toContain("folder/deep/deeper.txt");
    });

    it("should return empty array for empty ZIP", () => {
      const zip = new MiniZip();
      const data = zip.generate();
      const unzip = new MiniUnzip(data);

      const files = unzip.listFiles();
      expect(files).toEqual([]);
    });
  });
});

describe("MiniZip/MiniUnzip roundtrip", () => {
  it("should preserve content through zip/unzip cycle", () => {
    const testFiles = [
      { path: "text.txt", content: "Simple text" },
      { path: "numbers.txt", content: "123456789" },
      { path: "special.txt", content: "!@#$%^&*()" },
      { path: "nested/path/file.xml", content: "<root><child/></root>" }
    ];

    const zip = new MiniZip();
    testFiles.forEach(f => zip.addFile(f.path, f.content));

    const data = zip.generate();
    const unzip = new MiniUnzip(data);

    testFiles.forEach(f => {
      expect(unzip.getFile(f.path)).toBe(f.content);
    });
  });

  it("should handle XML content (XLSX use case)", () => {
    const xmlContent = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>`;

    const zip = new MiniZip();
    zip.addFile("xl/worksheets/sheet1.xml", xmlContent);

    const data = zip.generate();
    const unzip = new MiniUnzip(data);

    expect(unzip.getFile("xl/worksheets/sheet1.xml")).toBe(xmlContent);
  });

  it("should handle XLSX file structure", () => {
    const zip = new MiniZip();
    zip.addFile("[Content_Types].xml", '<?xml version="1.0"?><Types/>');
    zip.addFile("_rels/.rels", '<?xml version="1.0"?><Relationships/>');
    zip.addFile("xl/workbook.xml", '<?xml version="1.0"?><workbook/>');
    zip.addFile("xl/_rels/workbook.xml.rels", '<?xml version="1.0"?><Relationships/>');
    zip.addFile("xl/styles.xml", '<?xml version="1.0"?><styleSheet/>');
    zip.addFile("xl/sharedStrings.xml", '<?xml version="1.0"?><sst/>');
    zip.addFile("xl/worksheets/sheet1.xml", '<?xml version="1.0"?><worksheet/>');

    const data = zip.generate();
    const unzip = new MiniUnzip(data);

    const files = unzip.listFiles();
    expect(files).toContain("[Content_Types].xml");
    expect(files).toContain("xl/workbook.xml");
    expect(files).toContain("xl/worksheets/sheet1.xml");
  });
});

describe("MiniUnzip limits", () => {
  it("inflates a small deflate entry inside the output cap", async () => {
    const compressed = deflateRawSync(Buffer.from("hola"));
    const out = await MiniUnzip.inflateRaw(new Uint8Array(compressed), 1024);
    expect(new TextDecoder().decode(out)).toBe("hola");
  });

  it("rejects an entry that inflates past the output cap", async () => {
    const compressed = deflateRawSync(Buffer.alloc(50_000, 0x41));
    await expect(
      MiniUnzip.inflateRaw(new Uint8Array(compressed), 1024),
    ).rejects.toBeInstanceOf(UnzipLimitError);
  });

  it("does not inflate entries other than the one requested", async () => {
    const invalidDeflate = localFile("xl/bomb.xml", 8, new Uint8Array([0x00, 0x01, 0x02, 0x03]));
    const target = localFile(
      "xl/sharedStrings.xml",
      0,
      new TextEncoder().encode("<sst/>"),
    );
    const unzip = new MiniUnzip(concatBytes([invalidDeflate, target]));

    await expect(unzip.getFileAsync("xl/sharedStrings.xml")).resolves.toBe("<sst/>");
  });

  it("keeps the default entry limit and honors a lower one from the constructor", () => {
    const zip = new MiniZip();
    zip.addFile("a.txt", "x".repeat(200));
    const data = zip.generate();

    expect(new MiniUnzip(data).getFile("a.txt")).toBe("x".repeat(200));
    expect(() => new MiniUnzip(data, { maxEntryBytes: 100 }).getFile("a.txt")).toThrow(UnzipLimitError);
  });

  it("stops when a local header claims more bytes than the buffer", () => {
    const name = new TextEncoder().encode("a.txt");
    const out = new Uint8Array(30 + name.length + 4);
    const view = new DataView(out.buffer);
    view.setUint32(0, 0x04034b50, true);
    view.setUint32(18, 0xffffffff, true);
    view.setUint16(26, name.length, true);
    out.set(name, 30);

    const unzip = new MiniUnzip(out);
    expect(unzip.getFile("a.txt")).toBeNull();
    expect(unzip.listFiles()).toEqual([]);
  });
});
