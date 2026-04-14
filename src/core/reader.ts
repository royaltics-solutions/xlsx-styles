/**
 * Minimal XLSX Reader implementation
 */

import { MiniUnzip } from '../zip';

export class XlsxReader {
  private zip: MiniUnzip;
  
  constructor(data: Uint8Array) {
    this.zip = new MiniUnzip(data);
  }
  
  private parseSharedStrings(): string[] {
    const content = this.zip.getFile('xl/sharedStrings.xml');
    if (!content) return [];
    
    const strings: string[] = [];
    const regex = /<t[^>]*>(.*?)<\/t>/g;
    let match;
    
    while ((match = regex.exec(content)) !== null) {
      strings.push(this.unescapeXml(match[1]));
    }
    
    return strings;
  }
  
  private unescapeXml(str: string): string {
    return str
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, '&');
  }
  
  private parseWorksheetNames(): string[] {
    const content = this.zip.getFile('xl/workbook.xml');
    if (!content) return [];
    
    const names: string[] = [];
    const regex = /<sheet[^>]*name="([^"]*)"[^>]*>/g;
    let match;
    
    while ((match = regex.exec(content)) !== null) {
      names.push(this.unescapeXml(match[1]));
    }
    
    return names;
  }
  
  private parseWorksheet(sheetIndex: number, sharedStrings: string[]): any[][] {
    const content = this.zip.getFile(`xl/worksheets/sheet${sheetIndex + 1}.xml`);
    if (!content) return [];

    const rows: any[][] = [];
    const rowRegex = /<row[^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
    const cellRegex = /<c\s+r="([A-Z]+)(\d+)"(?:\s+t="([^"]*)")?[^>]*>[\s\S]*?<v>([^<]*)<\/v>/g;

    let rowMatch;
    while ((rowMatch = rowRegex.exec(content)) !== null) {
      const rowNum = parseInt(rowMatch[1]) - 1;
      const rowContent = rowMatch[2];
      const row: any[] = [];

      let cellMatch;
      cellRegex.lastIndex = 0;
      while ((cellMatch = cellRegex.exec(rowContent)) !== null) {
        const colLetters = cellMatch[1];
        const type = cellMatch[3];
        const value = cellMatch[4];

        let cellValue: any;

        if (type === 's') {
          cellValue = sharedStrings[parseInt(value)] || '';
        } else if (type === 'b') {
          cellValue = value === '1';
        } else if (type === 'str') {
          cellValue = this.unescapeXml(value);
        } else {
          const num = parseFloat(value);
          cellValue = isNaN(num) ? value : num;
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
  
  read(): { sheets: { name: string; data: any[][] }[] } {
    const sharedStrings = this.parseSharedStrings();
    const sheetNames = this.parseWorksheetNames();
    
    const sheets = sheetNames.map((name, i) => ({
      name,
      data: this.parseWorksheet(i, sharedStrings)
    }));
    
    return { sheets };
  }
}
