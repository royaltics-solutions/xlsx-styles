import { describe, it, expect, beforeAll, afterAll } from "vitest";
import { XlsxWriter, XlsxReader } from "../src/xlsx";
import { Xldx } from "../src/index";
import { readFile } from "../src/server";
import * as fs from "fs/promises";
import * as path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const TEST_OUTPUT_DIR = path.join(__dirname, "output");

beforeAll(async () => {
  await fs.mkdir(TEST_OUTPUT_DIR, { recursive: true });
});

afterAll(async () => {
  await fs.rm(TEST_OUTPUT_DIR, { recursive: true, force: true });
});

describe("XlsxReader - basic read", () => {
  it("should read back written data", async () => {
    const writer = new XlsxWriter();
    writer.addWorksheet("TestSheet", [
      ["Name", "Age"],
      ["Alice", 30],
      ["Bob", 25]
    ]);

    const xlsx = writer.generate();
    const reader = new XlsxReader(xlsx);
    const result = await reader.read();

    expect(result.sheets).toHaveLength(1);
    expect(result.sheets[0].name).toBe("TestSheet");
    expect(result.sheets[0].data[0]).toEqual(["Name", "Age"]);
    expect(result.sheets[0].data[1]).toEqual(["Alice", 30]);
    expect(result.sheets[0].data[2]).toEqual(["Bob", 25]);
  });

  it("should read multiple sheets", async () => {
    const writer = new XlsxWriter();
    writer.addWorksheet("Sheet1", [["A"], ["B"]]);
    writer.addWorksheet("Sheet2", [["C"], ["D"]]);

    const xlsx = writer.generate();
    const reader = new XlsxReader(xlsx);
    const result = await reader.read();

    expect(result.sheets).toHaveLength(2);
    expect(result.sheets[0].name).toBe("Sheet1");
    expect(result.sheets[1].name).toBe("Sheet2");
  });

  it("should handle numbers, booleans, and empty cells", async () => {
    const writer = new XlsxWriter();
    writer.addWorksheet("Mixed", [
      [1, 2.5, true, null],
      [false, 0, -3, "text"]
    ]);

    const xlsx = writer.generate();
    const reader = new XlsxReader(xlsx);
    const result = await reader.read();

    expect(result.sheets[0].data[0][0]).toBe(1);
    expect(result.sheets[0].data[0][1]).toBe(2.5);
    expect(result.sheets[0].data[0][2]).toBe(true);
    expect(result.sheets[0].data[1][0]).toBe(false);
    expect(result.sheets[0].data[1][3]).toBe("text");
  });
});

describe("XlsxReader - sheetToJson", () => {
  it("should convert to JSON with custom headers", () => {
    const data = [
      ["A1", "B1", "C1"],
      ["A2", "B2", "C2"]
    ];
    const reader = new XlsxReader(new Uint8Array(0));
    const json = reader.sheetToJson(data, {
      header: ["col_a", "col_b", "col_c"]
    });

    expect(json).toHaveLength(2);
    expect(json[0]).toEqual({ col_a: "A1", col_b: "B1", col_c: "C1" });
    expect(json[1]).toEqual({ col_a: "A2", col_b: "B2", col_c: "C2" });
  });

  it("should use first row as headers with header: 0", () => {
    const data = [
      ["Name", "Age"],
      ["Alice", 30],
      ["Bob", 25]
    ];
    const reader = new XlsxReader(new Uint8Array(0));
    const json = reader.sheetToJson(data, { header: 0 });

    expect(json).toHaveLength(2);
    expect(json[0]).toEqual({ Name: "Alice", Age: 30 });
    expect(json[1]).toEqual({ Name: "Bob", Age: 25 });
  });

  it("should use specified row index as headers", () => {
    const data = [
      ["ignore", "this"],
      ["Key1", "Key2"],
      ["val1", "val2"]
    ];
    const reader = new XlsxReader(new Uint8Array(0));
    const json = reader.sheetToJson(data, { header: 1 });

    expect(json).toHaveLength(1);
    expect(json[0]).toEqual({ Key1: "val1", Key2: "val2" });
  });

  it("should skip blank rows with blankrows option", () => {
    const data = [
      ["H1", "H2"],
      ["a", "b"],
      [null, null],
      ["c", "d"]
    ];
    const reader = new XlsxReader(new Uint8Array(0));
    const json = reader.sheetToJson(data, { header: 0, blankrows: true });

    expect(json).toHaveLength(2);
    expect(json[0]).toEqual({ H1: "a", H2: "b" });
    expect(json[1]).toEqual({ H1: "c", H2: "d" });
  });

  it("should use defval for empty cells", () => {
    const data = [
      ["H1", "H2"],
      ["a", null]
    ];
    const reader = new XlsxReader(new Uint8Array(0));
    const json = reader.sheetToJson(data, { header: 0, defval: "N/A" });

    expect(json[0]).toEqual({ H1: "a", H2: "N/A" });
  });
});

describe("Xldx.read with options", () => {
  it("should read and return json output", async () => {
    const xldx = new Xldx([
      { name: "Alice", age: 30 },
      { name: "Bob", age: 25 }
    ]);
    xldx.createSheet(
      { name: "People" },
      { key: "name", header: "Name" },
      { key: "age", header: "Age" }
    );

    const uint8Array = await xldx.toUint8Array();
    const result = await Xldx.read(uint8Array, {
      header: ["full_name", "years"]
    });

    expect(result.sheets[0].json).toBeDefined();
    // With custom headers, every row is mapped (including header row)
    expect(result.sheets[0].json).toHaveLength(3);
    expect(result.sheets[0].json[1]).toEqual({ full_name: "Alice", years: 30 });
    expect(result.sheets[0].json[2]).toEqual({ full_name: "Bob", years: 25 });
  });
});

describe("readFile - Node.js integration", () => {
  it("should read an XLSX file from disk", async () => {
    const data = [{ name: "Test", value: 42 }];
    const xldx = new Xldx(data);
    xldx.createSheet(
      { name: "ReadFileTest" },
      { key: "name", header: "Name" },
      { key: "value", header: "Value" }
    );

    const filePath = path.join(TEST_OUTPUT_DIR, "readfile-test.xlsx");
    await xldx.write(filePath);

    const result = await readFile(filePath);
    expect(result.sheets).toHaveLength(1);
    expect(result.sheets[0].name).toBe("ReadFileTest");
    expect(result.sheets[0].data[0]).toEqual(["Name", "Value"]);
    expect(result.sheets[0].data[1]).toEqual(["Test", 42]);
  });

  it("should readFile with json output using options", async () => {
    const data = [
      { name: "Alice", age: 30 },
      { name: "Bob", age: 25 }
    ];

    const xldx = new Xldx(data);
    xldx.createSheet(
      { name: "People" },
      { key: "name", header: "Name" },
      { key: "age", header: "Age" }
    );

    const filePath = path.join(TEST_OUTPUT_DIR, "readfile-json.xlsx");
    await xldx.write(filePath);

    const result = await readFile(filePath, {
      header: ["full_name", "years"],
      blankrows: true
    });

    expect(result.sheets[0].json).toBeDefined();
    expect(result.sheets[0].json).toHaveLength(3);
    expect(result.sheets[0].json[1]).toEqual({ full_name: "Alice", years: 30 });
    expect(result.sheets[0].json[2]).toEqual({ full_name: "Bob", years: 25 });
  });
});

describe("Real file test.xlsx from disk", () => {
  const fixturePath = path.join(__dirname, "test.xlsx");

  it("should read a real .xlsx file from disk (place your own test.xlsx in test/)", async () => {
    try {
      await fs.access(fixturePath);
    } catch {
      return;
    }
    const result = await readFile(fixturePath, { header: ['A','B','C','D','E','F']});

    result.sheets.forEach((sheet, i) => {
      console.log(`\n--- Sheet ${i}: "${sheet.name}" (${sheet.data.length} rows) ---`);
      console.log("  Header row:", JSON.stringify(sheet.data[0]));
      if (sheet.data.length > 1) {
        console.log("  First data row:", JSON.stringify(sheet.data[1]));
      }
    });

    expect(result.sheets.length).toBeGreaterThan(0);
    expect(result.sheets[0].data.length).toBeGreaterThan(0);
  });

  it("should read test.xlsx with custom headers via readFile", async () => {
    try {
      await fs.access(fixturePath);
    } catch {
      return;
    }
    const result = await readFile(fixturePath, {
      header: ["A", "B", "C", "D", "E", "F"],
    });

    result.sheets.forEach((sheet, i) => {
      if (sheet.json) {
        console.log(`\n--- JSON Sheet ${i}: "${sheet.name}" ---`);
        console.log(JSON.stringify(sheet.json.slice(0, 5), null, 2));
      }
    });

    expect(result.sheets[0].json).toBeDefined();
  });
});

describe("End-to-end: simulate CFDI_DOWNLOAD.IMPORT_XLSX pattern", () => {
  it("should import Excel, map with custom headers, transform rows", async () => {
    // Create a file similar to what the user imports
    const xldx = new Xldx([]);
    xldx.createSheet(
      { name: "CFDI" },
      { key: "document_type", header: "document_type" },
      { key: "datedoc", header: "datedoc" },
      { key: "cardid", header: "cardid" },
      { key: "keyaccess", header: "keyaccess" }
    );
    
    // Manually add data matching the user's import format
    // keyaccess format: DDMMYYYY+type(2)+pad(14)+serial1(3)+serial2(3)+serial3(9)+pad(10) = 49 chars
    const writer = new XlsxWriter();
    writer.addWorksheet("Sheet1", [
      ["01", "01012024", "CARD001", "0101202401001001000000001123456789012345678901234"],
      ["01", "02012024", "CARD002", "0201202401001001000000002234567890123456789012345"],
      ["02", "03012024", "CARD003", "0301202402001001000000003234567890123456789012345"],
    ]);

    const xlsx = writer.generate();
    const result = await Xldx.read(xlsx, {
      header: ["document_type", "datedoc", "cardid", "keyaccess", "sri_support", "as", "code_tax"],
      blankrows: false,
    });

    expect(result.sheets[0].json).toBeDefined();

    // Process rows like the user's example
    const processed: any[] = [];
    for (let data of result.sheets[0].json!) {
      data.keyaccess = String(data.keyaccess || '').replace(/\'/gi, "").replace(/\s/gi, "").substring(0, 49);
      if (data.keyaccess?.length != 49) continue;

      data.datedoc = data.keyaccess.substring(4, 8) + '-' + data.keyaccess.substring(2, 4) + "-" + data.keyaccess.substring(0, 2);
      const number = data.keyaccess.substring(24, 27) + '-' + data.keyaccess.substring(27, 30) + '-' + data.keyaccess.substring(30, 39);

      processed.push({
        keyaccess: data.keyaccess,
        datedoc: data.datedoc,
        number,
        document_type: data.document_type,
      });
    }

    expect(processed).toHaveLength(3);
    expect(processed[0].datedoc).toBe("2024-01-01");
    expect(processed[0].number).toBe("112-345-678901234");
    expect(processed[1].datedoc).toBe("2024-01-02");
  });
});

describe("Writer is untouched", () => {
  it("should still generate styled XLSX", async () => {
    const data = [
      { name: "Alice", age: 30, city: "New York" },
      { name: "Bob", age: 25, city: "Los Angeles" },
    ];

    const xldx = new Xldx(data);
    xldx.createSheet(
      { name: "People" },
      { key: "name", header: "Name", width: "auto" },
      { key: "age", header: "Age" },
      { key: "city", header: "City" }
    );

    const buffer = await xldx.toBuffer();
    expect(buffer).toBeInstanceOf(Buffer);
    expect(buffer.length).toBeGreaterThan(0);
    expect(buffer[0]).toBe(0x50); // PK
  });

  it("should support themes and patterns", async () => {
    const data = Array.from({ length: 10 }, (_, i) => ({
      category: String.fromCharCode(65 + i),
      value: i * 10,
    }));

    const xldx = new Xldx(data);
    xldx.createSheet(
      { name: "Themed" },
      { key: "category", header: "Category", patterns: { bgColorPattern: "colorPerDiff" } },
      { key: "value", header: "Value", patterns: { bgColorPattern: "zebra" } }
    );

    const uint8Array = await xldx.toUint8Array();
    expect(uint8Array.length).toBeGreaterThan(0);
  });

  it("should produce same output format (read roundtrip)", async () => {
    const originalData = [
      ["Header1", "Header2"],
      ["Data1", 100],
    ];

    const writer = new XlsxWriter();
    writer.addWorksheet("Test", originalData);

    const xlsx = writer.generate();
    const reader = new XlsxReader(xlsx);
    const result = await reader.read();

    expect(result.sheets[0].data[0]).toEqual(["Header1", "Header2"]);
    expect(result.sheets[0].data[1]).toEqual(["Data1", 100]);
  });
});
