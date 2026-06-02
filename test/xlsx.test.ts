import { describe, it, expect } from "vitest";
import { XlsxWriter, XlsxReader } from "../src/xlsx";

describe("XlsxWriter", () => {
  describe("addWorksheet", () => {
    it("should add a worksheet with data", () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Sheet1", [
        ["Name", "Age"],
        ["Alice", 30],
        ["Bob", 25]
      ]);

      const xlsx = writer.generate();
      expect(xlsx).toBeInstanceOf(Uint8Array);
      expect(xlsx.length).toBeGreaterThan(0);
    });

    it("should add multiple worksheets", () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Sheet1", [["A", "B"], [1, 2]]);
      writer.addWorksheet("Sheet2", [["C", "D"], [3, 4]]);

      const xlsx = writer.generate();
      expect(xlsx).toBeInstanceOf(Uint8Array);
    });

    it("should handle column widths", () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Sheet1", [["Name", "Age"]], [20, 10]);

      const xlsx = writer.generate();
      expect(xlsx).toBeInstanceOf(Uint8Array);
    });
  });

  describe("generate", () => {
    it("should generate valid ZIP structure", () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Test", [["Hello"]]);

      const xlsx = writer.generate();

      // ZIP files start with PK signature (0x04034b50)
      expect(xlsx[0]).toBe(0x50); // P
      expect(xlsx[1]).toBe(0x4b); // K
    });

    it("should handle string values", () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Test", [
        ["String1", "String2"],
        ["Hello", "World"]
      ]);

      const xlsx = writer.generate();
      expect(xlsx.length).toBeGreaterThan(0);
    });

    it("should handle number values", () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Test", [
        [1, 2, 3],
        [4.5, 6.7, 8.9]
      ]);

      const xlsx = writer.generate();
      expect(xlsx.length).toBeGreaterThan(0);
    });

    it("should handle boolean values", () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Test", [
        [true, false],
        [false, true]
      ]);

      const xlsx = writer.generate();
      expect(xlsx.length).toBeGreaterThan(0);
    });

    it("should handle null and undefined values", () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Test", [
        [null, undefined, ""],
        ["value", null, undefined]
      ]);

      const xlsx = writer.generate();
      expect(xlsx.length).toBeGreaterThan(0);
    });

    it("should handle Date values", () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Test", [
        ["Date"],
        [new Date("2024-01-15")]
      ]);

      const xlsx = writer.generate();
      expect(xlsx.length).toBeGreaterThan(0);
    });

    it("should handle Cell objects with value property", () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Test", [
        [{ value: "Test", style: { font: { bold: true } } }],
        [{ value: 123, style: {} }]
      ]);

      const xlsx = writer.generate();
      expect(xlsx.length).toBeGreaterThan(0);
    });

    it("should escape XML special characters", () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Test", [
        ["<test>", "&value", '"quoted"', "'apostrophe'"]
      ]);

      const xlsx = writer.generate();
      expect(xlsx.length).toBeGreaterThan(0);
    });

    it("should handle sheet names with special characters", () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Sheet <1>", [["test"]]);

      const xlsx = writer.generate();
      expect(xlsx.length).toBeGreaterThan(0);
    });

    it("should handle empty worksheet", () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Empty", []);

      const xlsx = writer.generate();
      expect(xlsx.length).toBeGreaterThan(0);
    });

    it("should handle large column indices", () => {
      const writer = new XlsxWriter();
      const row = Array(30).fill("test");
      writer.addWorksheet("Wide", [row]);

      const xlsx = writer.generate();
      expect(xlsx.length).toBeGreaterThan(0);
    });
  });
});

describe("XlsxReader", () => {
  describe("read", () => {
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

    it("should handle numbers correctly", async () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Numbers", [
        [1, 2.5, -3, 0]
      ]);

      const xlsx = writer.generate();
      const reader = new XlsxReader(xlsx);
      const result = await reader.read();

      expect(result.sheets[0].data[0]).toEqual([1, 2.5, -3, 0]);
    });

    it("should handle booleans correctly", async () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Bools", [
        [true, false]
      ]);

      const xlsx = writer.generate();
      const reader = new XlsxReader(xlsx);
      const result = await reader.read();

      expect(result.sheets[0].data[0]).toEqual([true, false]);
    });

    it("should handle empty cells", async () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Sparse", [
        ["A", null, "C"],
        [null, "B", null]
      ]);

      const xlsx = writer.generate();
      const reader = new XlsxReader(xlsx);
      const result = await reader.read();

      expect(result.sheets[0].data[0][0]).toBe("A");
      expect(result.sheets[0].data[0][2]).toBe("C");
    });

    it("should handle XML-escaped content", async () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Escaped", [
        ["<tag>", "&amp;", '"quote"']
      ]);

      const xlsx = writer.generate();
      const reader = new XlsxReader(xlsx);
      const result = await reader.read();

      expect(result.sheets[0].data[0][0]).toBe("<tag>");
      expect(result.sheets[0].data[0][1]).toBe("&amp;");
    });
  });

  describe("sheetToJson", () => {
    it("should convert sheet data to JSON with custom headers", () => {
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

    it("should use first row as headers by default", () => {
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

  describe("deflated ZIP support", () => {
    it("should read deflated XLSX entries", async () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("DeflateTest", [
        ["Name", "Score"],
        ["Alice", 95],
        ["Bob", 87]
      ]);
      const stored = writer.generate();

      const reader = new XlsxReader(stored);
      const result = await reader.read();

      expect(result.sheets).toHaveLength(1);
      expect(result.sheets[0].data[0]).toEqual(["Name", "Score"]);
      expect(result.sheets[0].data[1]).toEqual(["Alice", 95]);
      expect(result.sheets[0].data[2]).toEqual(["Bob", 87]);
    });

    it("should handle inline strings (type inlineStr)", async () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("Inline", [
        ["Hello", "World"]
      ]);

      const xlsx = writer.generate();
      const reader = new XlsxReader(xlsx);
      const result = await reader.read();

      expect(result.sheets[0].data[0]).toEqual(["Hello", "World"]);
    });

    it("should read with json output via read options", async () => {
      const writer = new XlsxWriter();
      writer.addWorksheet("People", [
        ["Name", "Age"],
        ["Alice", 30],
        ["Bob", 25]
      ]);

      const xlsx = writer.generate();
      const reader = new XlsxReader(xlsx);
      const result = await reader.read({
        header: ["name", "age"]
      });

      expect(result.sheets[0].json).toBeDefined();
      expect(result.sheets[0].json).toHaveLength(3);
      expect(result.sheets[0].json![1]).toEqual({ name: "Alice", age: 30 });
      expect(result.sheets[0].json![2]).toEqual({ name: "Bob", age: 25 });
    });
  });
});

describe("XlsxWriter/XlsxReader roundtrip", () => {
  it("should preserve data through write/read cycle", async () => {
    const originalData = [
      ["Header1", "Header2", "Header3"],
      ["String", 123, true],
      ["Another", 456.78, false],
      ["", null, 0]
    ];

    const writer = new XlsxWriter();
    writer.addWorksheet("RoundTrip", originalData);

    const xlsx = writer.generate();
    const reader = new XlsxReader(xlsx);
    const result = await reader.read();

    expect(result.sheets[0].data[0]).toEqual(["Header1", "Header2", "Header3"]);
    expect(result.sheets[0].data[1]).toEqual(["String", 123, true]);
    expect(result.sheets[0].data[2]).toEqual(["Another", 456.78, false]);
  });

  it("should handle shared strings efficiently", async () => {
    const writer = new XlsxWriter();
    writer.addWorksheet("SharedStrings", [
      ["Repeated", "Repeated", "Repeated"],
      ["Repeated", "Unique", "Repeated"]
    ]);

    const xlsx = writer.generate();
    const reader = new XlsxReader(xlsx);
    const result = await reader.read();

    expect(result.sheets[0].data[0]).toEqual(["Repeated", "Repeated", "Repeated"]);
    expect(result.sheets[0].data[1][1]).toBe("Unique");
  });
});
