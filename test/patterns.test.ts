import { describe, it, expect } from "vitest";
import {
  zebraBg,
  bgColorBasedOnDiff,
  txtColorBasedOnDiff,
  createSetWidthBasedOnCharacterCount,
  applyPattern,
  buildPatternContext
} from "../src/utils";
import { defaultTheme } from "../src/themes";
import type { PatternContext, DataRow } from "../src";

describe("Pattern Functions", () => {
  it("should apply zebra background to even rows", () => {
    const context: PatternContext = {
      rowIndex: 2,
      columnIndex: 0,
      value: 'test',
      rowData: { col1: 'test' },
      allData: [{ col1: 'test' }],
      columnKey: 'col1'
    };
    
    const result = zebraBg(context);
    expect(result).toEqual({
      fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: defaultTheme.base[100]
      }
    });
  });

  it("should return null for odd zebra rows", () => {
    const context: PatternContext = {
      rowIndex: 3,
      columnIndex: 0,
      value: 'test',
      rowData: { col1: 'test' },
      allData: [{ col1: 'test' }],
      columnKey: 'col1'
    };
    
    const result = zebraBg(context);
    expect(result).toBeNull();
  });

  it("should assign different colors to different values", () => {
    const allData: DataRow[] = [
      { category: 'A' },
      { category: 'B' },
      { category: 'A' },
      { category: 'C' }
    ];

    const contextA: PatternContext = {
      rowIndex: 1,
      columnIndex: 0,
      value: 'A',
      rowData: { category: 'A' },
      allData,
      columnKey: 'category'
    };

    const contextB: PatternContext = {
      rowIndex: 2,
      columnIndex: 0,
      value: 'B',
      rowData: { category: 'B' },
      allData,
      columnKey: 'category'
    };

    const resultA = bgColorBasedOnDiff(contextA);
    const resultB = bgColorBasedOnDiff(contextB);

    expect(resultA?.fill?.fgColor).toBeDefined();
    expect(resultB?.fill?.fgColor).toBeDefined();
    expect(resultA?.fill?.fgColor).not.toEqual(resultB?.fill?.fgColor);
  });

  it("should highlight changed text values", () => {
    const context: PatternContext = {
      rowIndex: 2,
      columnIndex: 0,
      value: 'new',
      previousValue: 'old',
      rowData: { col1: 'new' },
      allData: [{ col1: 'old' }, { col1: 'new' }],
      columnKey: 'col1'
    };

    const result = txtColorBasedOnDiff(context);
    expect(result).toEqual({
      font: {
        color: defaultTheme.primary,
        bold: true
      }
    });
  });

  it("should calculate column width from data", () => {
    const columnData = ['short', 'medium text', 'long text with more characters'];
    const calculator = createSetWidthBasedOnCharacterCount(columnData);
    const result = calculator();

    expect(result).toBeDefined();
    expect(result?.width).toBeGreaterThan(10);
    expect(result?.wrapText).toBe(true);
  });

  it("should apply patterns when creating sheets", async () => {
    const data = [
      { name: 'Alice', score: 95 },
      { name: 'Bob', score: 87 },
      { name: 'Charlie', score: 95 }
    ];

    const xldx = new (await import("../src/server")).Xldx(data);
    
    xldx.createSheet(
      { name: 'Scores' },
      {
        key: 'name',
        header: 'Name',
        patterns: {
          bgColorPattern: 'zebra'
        }
      },
      {
        key: 'score',
        header: 'Score',
        patterns: {
          bgColorPattern: 'colorPerDiff'
        }
      }
    );

    const sheetData = xldx.getSheetData(0);
    const rows = sheetData.getRowsData();
    
    expect(rows).toHaveLength(3);
    expect(rows[0]).toEqual({ name: 'Alice', score: 95 });
  });

  it("should apply pattern by name or function", () => {
    const context: PatternContext = {
      rowIndex: 2,
      columnIndex: 0,
      value: 'test',
      rowData: { col1: 'test' },
      allData: [{ col1: 'test' }],
      columnKey: 'col1'
    };

    const result = applyPattern('zebra', context);
    expect(result).toBeDefined();

    const customPattern = (ctx: PatternContext) => ({
      font: { color: '#FF0000' }
    });

    const customResult = applyPattern(customPattern, context);
    expect(customResult).toEqual({ font: { color: '#FF0000' } });
  });

  it("should build pattern context with proper offsets", () => {
    const params = {
      rowIndex: 0,
      colIndex: 1,
      rowData: { col1: 'A', col2: 'B' },
      columnKey: 'col2',
      value: 'B',
      allData: [{ col1: 'A', col2: 'B' }]
    };

    const context = buildPatternContext(params);
    expect(context.rowIndex).toBe(2);
    expect(context.columnIndex).toBe(1);
    expect(context.value).toBe('B');
    expect(context.columnKey).toBe('col2');
  });
});
