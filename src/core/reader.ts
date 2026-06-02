/**
 * XLSX Reader implementation
 * Supports both stored and deflated ZIP entries,
 * shared strings (including rich text), and sheet-to-JSON conversion.
 */

import { MiniUnzip } from '../zip';

export interface ReadOptions {
  /** Column headers. 
   *  string[]: use as header keys, map by column position.
   *  number: row index to use as headers (0-based).
   *  undefined: use first non-empty row as headers. */
  header?: string[] | number;
  /** Skip completely empty rows when generating json (default: false). */
  blankrows?: boolean;
  /** Row range to process. If number, reads rows [0..range). 
   *  If object, reads rows in [{s:{r,c}, e:{r,c}}]. */
  range?: number | { s: { r: number; c: number }; e: { r: number; c: number } };
  /** Default value for empty/undefined cells (default: null). */
  defval?: any;
  /** Return raw values (default: true). false attempts formatted strings. */
  raw?: boolean;
  /** Maximum number of rows to parse per sheet. */
  sheetRows?: number;
}

export interface SheetResult {
  name: string;
  data: any[][];
  json?: Record<string, any>[];
}

function columnLetter(col: number): string {
  let s = '';
  let n = col;
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

export class XlsxReader {
  private zip: MiniUnzip;
  
  constructor(data: Uint8Array) {
    this.zip = new MiniUnzip(data);
  }
  
  private unescapeXml(str: string): string {
    return str
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, '&');
  }
  
  private async loadFile(path: string): Promise<string | null> {
    return this.zip.getFileAsync(path);
  }
  
  private async parseSharedStrings(): Promise<string[]> {
    const content = await this.loadFile('xl/sharedStrings.xml');
    if (!content) return [];
    
    const strings: string[] = [];
    const siRegex = /<si>(.*?)<\/si>/g;
    let siMatch;
    
    while ((siMatch = siRegex.exec(content)) !== null) {
      const siContent = siMatch[1];
      const tRegex = /<t[^>]*>(.*?)<\/t>/g;
      let tMatch;
      let text = '';
      while ((tMatch = tRegex.exec(siContent)) !== null) {
        text += this.unescapeXml(tMatch[1]);
      }
      strings.push(text);
    }
    
    return strings;
  }
  
  private async parseWorksheetNames(): Promise<string[]> {
    const content = await this.loadFile('xl/workbook.xml');
    if (!content) return [];
    
    const names: string[] = [];
    const regex = /<sheet[^>]*name="([^"]*)"[^>]*>/g;
    let match;
    
    while ((match = regex.exec(content)) !== null) {
      names.push(this.unescapeXml(match[1]));
    }
    
    return names;
  }
  
  private async parseWorksheet(sheetIndex: number, sharedStrings: string[]): Promise<any[][]> {
    const content = await this.loadFile(`xl/worksheets/sheet${sheetIndex + 1}.xml`);
    if (!content) return [];

    const rows: any[][] = [];
    const rowRegex = /<row[^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
    const cellRegex = /<c\s+r="([A-Z]+)(\d+)"(?:\s+t="([^"]*)")?[^>]*>([\s\S]*?)<\/c>/g;

    let rowMatch;
    while ((rowMatch = rowRegex.exec(content)) !== null) {
      const rowNum = parseInt(rowMatch[1]) - 1;
      const rowContent = rowMatch[2];

      if (rowContent.trim().length === 0) continue;

      const row: any[] = [];
      let cellMatch;
      cellRegex.lastIndex = 0;
      
      while ((cellMatch = cellRegex.exec(rowContent)) !== null) {
        const colLetters = cellMatch[1];
        const fullTag = cellMatch[0];
        const typeMatch = fullTag.match(/\s+t="([^"]+)"/);
        const type = typeMatch ? typeMatch[1] : 'n';
        const cellContent = cellMatch[4];

        let cellValue: any;

        if (type === 's') {
          const vMatch = /<v>([^<]*)<\/v>/.exec(cellContent);
          cellValue = vMatch ? sharedStrings[parseInt(vMatch[1])] || '' : '';
        } else if (type === 'inlineStr' || type === 'str') {
          const tMatch = /<t[^>]*>(.*?)<\/t>/.exec(cellContent);
          cellValue = tMatch ? this.unescapeXml(tMatch[1]) : '';
        } else if (type === 'b') {
          const vMatch = /<v>([^<]*)<\/v>/.exec(cellContent);
          cellValue = vMatch ? vMatch[1] === '1' : false;
        } else {
          const vMatch = /<v>([^<]*)<\/v>/.exec(cellContent);
          const raw = vMatch ? vMatch[1] : '';
          const num = parseFloat(raw);
          cellValue = isNaN(num) ? raw : num;
        }

        const colIndex = this.lettersToColumn(colLetters);
        row[colIndex] = cellValue;
      }

      rows[rowNum] = row;
    }

    return rows;
  }
  
  private lettersToColumn(letters: string): number {
    let col = 0;
    for (let i = 0; i < letters.length; i++) {
      col = col * 26 + (letters.charCodeAt(i) - 64);
    }
    return col - 1;
  }
  
  sheetToJson(data: any[][], options: ReadOptions): Record<string, any>[] {
    const { blankrows, defval = null, range, sheetRows } = options;
    
    let headers: string[];
    let startRow: number;
    
    if (Array.isArray(options.header)) {
      headers = options.header;
      startRow = 0;
    } else if (typeof options.header === 'number') {
      const headerRow = data[options.header];
      headers = headerRow ? headerRow.map(String) : [];
      startRow = options.header + 1;
    } else {
      const maxCols = data.reduce((max, row) => Math.max(max, row?.length || 0), 0);
      headers = Array.from({ length: maxCols }, (_, i) => columnLetter(i));
      startRow = 0;
    }
    
    const result: Record<string, any>[] = [];
    
    let maxRow = data.length;
    if (typeof range === 'number') {
      maxRow = Math.min(maxRow, range);
    } else if (range && typeof range === 'object') {
      maxRow = Math.min(maxRow, (range as any).e?.r != null ? (range as any).e.r + 1 : maxRow);
      startRow = Math.max(startRow, (range as any).s?.r || 0);
    }
    if (sheetRows) {
      maxRow = Math.min(maxRow, startRow + sheetRows);
    }
    
    for (let i = startRow; i < maxRow; i++) {
      const row = data[i] || [];
      const obj: Record<string, any> = {};
      let hasValues = false;
      
      for (let j = 0; j < headers.length; j++) {
        const val = j < row.length ? row[j] : undefined;
        obj[headers[j]] = val !== undefined && val !== null ? val : defval;
        if (val !== undefined && val !== null && val !== '') {
          hasValues = true;
        }
      }
      
      if (blankrows && !hasValues) continue;
      result.push(obj);
    }
    
    return result;
  }
  
  async read(options?: ReadOptions): Promise<{ sheets: SheetResult[] }> {
    const sharedStrings = await this.parseSharedStrings();
    const sheetNames = await this.parseWorksheetNames();
    
    const sheets: SheetResult[] = [];
    for (let i = 0; i < sheetNames.length; i++) {
      const data = await this.parseWorksheet(i, sharedStrings);
      const sheet: SheetResult = { name: sheetNames[i], data, json: this.sheetToJson(data, options || {}) };
      sheets.push(sheet);
    }
    
    return { sheets };
  }
}
