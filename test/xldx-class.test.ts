import { describe, it, expect } from "vitest";
import { Xldx } from "../src/server";
import { defaultTheme } from "../src/themes";
import type { ColorTheme } from "../src";

describe("Xldx class methods", () => {
  describe("setTheme", () => {
    it("should set a custom theme and return this for chaining", () => {
      const customTheme: ColorTheme = {
        ...defaultTheme,
        primary: "#FF0000"
      };

      const xldx = new Xldx([{ a: 1 }]);
      const result = xldx.setTheme(customTheme);

      expect(result).toBe(xldx);
    });
  });

  describe("createColumn", () => {
    it("should return the column definition unchanged", () => {
      const xldx = new Xldx([]);
      const definition = { key: "test", header: "Test Header", width: 100 };

      const result = xldx.createColumn(definition);
      expect(result).toEqual(definition);
    });
  });

  describe("createColumns", () => {
    it("should return the column definitions unchanged", () => {
      const xldx = new Xldx([]);
      const definitions = [
        { key: "col1", header: "Column 1" },
        { key: "col2", header: "Column 2" }
      ];

      const result = xldx.createColumns(definitions);
      expect(result).toEqual(definitions);
    });
  });

  describe("createSheets", () => {
    it("should create multiple sheets at once", () => {
      const data = [
        { name: "Alice", age: 30 },
        { name: "Bob", age: 25 }
      ];

      const xldx = new Xldx(data);
      xldx.createSheets([
        {
          options: { name: "Sheet1" },
          columns: [{ key: "name", header: "Name" }]
        },
        {
          options: { name: "Sheet2" },
          columns: [{ key: "age", header: "Age" }]
        }
      ]);

      const sheet1 = xldx.getSheetData("Sheet1");
      const sheet2 = xldx.getSheetData("Sheet2");

      expect(sheet1.getRowsData()).toEqual(data);
      expect(sheet2.getRowsData()).toEqual(data);
    });
  });

  describe("toJSON", () => {
    it("should export workbook as JSON", () => {
      const data = [
        { name: "Alice", score: 95 },
        { name: "Bob", score: 87 }
      ];

      const xldx = new Xldx(data);
      xldx.createSheet(
        { name: "Scores" },
        { key: "name", header: "Name" },
        { key: "score", header: "Score" }
      );

      const json = xldx.toJSON();

      expect(json.sheets).toHaveLength(1);
      expect(json.sheets[0].name).toBe("Scores");
      expect(json.sheets[0].data).toBeDefined();
    });

    it("should export multiple sheets", () => {
      const xldx = new Xldx([{ a: 1 }]);
      xldx.createSheet({ name: "Sheet1" }, { key: "a" });

      const xldx2 = new Xldx([{ b: 2 }]);
      xldx2.createSheet({ name: "Sheet2" }, { key: "b" });

      const json1 = xldx.toJSON();
      const json2 = xldx2.toJSON();

      expect(json1.sheets).toHaveLength(1);
      expect(json2.sheets).toHaveLength(1);
    });
  });

  describe("fromJSON", () => {
    it("should create Xldx instance from JSON", () => {
      const json = {
        sheets: [
          {
            name: "TestSheet",
            data: [["Header"], ["Value"]],
            columnWidths: [20]
          }
        ]
      };

      const xldx = Xldx.fromJSON(json);
      expect(xldx).toBeInstanceOf(Xldx);
    });

    it("should handle empty sheets array", () => {
      const json = { sheets: [] };
      const xldx = Xldx.fromJSON(json);
      expect(xldx).toBeInstanceOf(Xldx);
    });

    it("should handle missing sheets property", () => {
      const json = {};
      const xldx = Xldx.fromJSON(json);
      expect(xldx).toBeInstanceOf(Xldx);
    });
  });
});
