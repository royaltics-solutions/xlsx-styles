import { MiniZip } from '../zip';
import type { Cell, Worksheet } from './types';
import type { CellStyle } from '../schemas';

export class XlsxWriter {
  private worksheets: Worksheet[] = [];
  private sharedStrings: string[] = [];
  private stringMap: Map<string, number> = new Map();

  private fonts: string[] = ['<font><sz val="11"/><name val="Calibri"/></font>'];
  private fills: string[] = [
    '<fill><patternFill patternType="none"/></fill>',
    '<fill><patternFill patternType="gray125"/></fill>'
  ];
  private numFmts: string[] = [];
  private cellXfs: string[] = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  private styleCache: Map<string, number> = new Map();

  addWorksheet(name: string, data: any[][], columnWidths?: number[], merges?: string[]): void {
    const cells: (Cell | string | number | boolean | null)[][] = data.map(row =>
      row.map(value => {
        if (typeof value === 'object' && value !== null && !(value instanceof Date) && 'value' in value) {
          return value as Cell;
        }
        return value;
      })
    );
    this.worksheets.push({ name, data: cells, columnWidths, merges });
  }

  private addSharedString(str: string): number {
    if (this.stringMap.has(str)) return this.stringMap.get(str)!;
    const index = this.sharedStrings.length;
    this.sharedStrings.push(str);
    this.stringMap.set(str, index);
    return index;
  }

  private normalizeColor(color?: string): string {
    if (!color) return 'FFFFFFFF';
    let c = color.replace('#', '').toUpperCase();
    if (c.length === 6) c = 'FF' + c;
    return c;
  }

  private getStyleId(style?: CellStyle): number {
    if (!style) return 0;
    const json = JSON.stringify(style);
    if (this.styleCache.has(json)) return this.styleCache.get(json)!;

    let fontId = 0;
    if (style.font) {
      const colorAttr = style.font.color ? `<color rgb="${this.normalizeColor(style.font.color)}"/>` : '';
      const fontXml = `<font>${style.font.bold ? '<b/>' : ''}${style.font.italic ? '<i/>' : ''}${colorAttr}<sz val="${style.font.size || 11}"/><name val="${style.font.name || 'Calibri'}"/></font>`;
      fontId = this.fonts.indexOf(fontXml);
      if (fontId === -1) {
        fontId = this.fonts.length;
        this.fonts.push(fontXml);
      }
    }

    let fillId = 0;
    if (style.fill?.fgColor) {
      const fillXml = `<fill><patternFill patternType="solid"><fgColor rgb="${this.normalizeColor(style.fill.fgColor)}"/><bgColor indexed="64"/></patternFill></fill>`;
      fillId = this.fills.indexOf(fillXml);
      if (fillId === -1) {
        fillId = this.fills.length;
        this.fills.push(fillXml);
      }
    }

    let numFmtId = 0;
    if (style.numFmt) {
      const builtIn: Record<string, number> = { 'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4 };
      if (builtIn[style.numFmt] !== undefined) {
        numFmtId = builtIn[style.numFmt];
      } else {
        numFmtId = 164 + this.numFmts.length;
        this.numFmts.push(`<numFmt numFmtId="${numFmtId}" formatCode="${this.escapeXml(style.numFmt)}"/>`);
      }
    }

    const xfXml = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="0" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1"/>`;
    let xfId = this.cellXfs.indexOf(xfXml);
    if (xfId === -1) {
      xfId = this.cellXfs.length;
      this.cellXfs.push(xfXml);
    }

    this.styleCache.set(json, xfId);
    return xfId;
  }

  private getCellValue(cell: any): { type: string; value: string; styleId: number } {
    const isObj = typeof cell === 'object' && cell !== null && !(cell instanceof Date) && 'value' in cell;
    const val = isObj ? cell.value : cell;
    const styleId = isObj ? this.getStyleId(cell.style) : 0;

    if (val === null || val === undefined) return { type: '', value: '', styleId };
    if (typeof val === 'boolean') return { type: 'b', value: val ? '1' : '0', styleId };
    if (typeof val === 'number') return { type: 'n', value: val.toString(), styleId };
    if (val instanceof Date) {
      const excelDate = Math.floor((val.getTime() - new Date(1900, 0, 1).getTime()) / 86400000) + 2;
      return { type: 'n', value: excelDate.toString(), styleId };
    }
    return { type: 's', value: this.addSharedString(String(val)).toString(), styleId };
  }

  private columnToLetters(col: number): string {
    let letters = '';
    col++;
    while (col > 0) {
      col--;
      letters = String.fromCharCode(65 + (col % 26)) + letters;
      col = Math.floor(col / 26);
    }
    return letters;
  }

  private generateContentTypes(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${this.worksheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`;
  }

  private generateRels(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  }

  private generateWorkbook(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${this.worksheets.map((s, i) => `<sheet name="${this.escapeXml(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`;
  }

  private generateWorkbookRels(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${this.worksheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${this.worksheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId${this.worksheets.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`;
  }

  private generateWorksheet(sheet: Worksheet): string {
    const cols = sheet.columnWidths ? `<cols>${sheet.columnWidths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>` : '';
    const rows = sheet.data.map((row, rIdx) => {
      const cells = row.map((cell, cIdx) => {
        const { type, value, styleId } = this.getCellValue(cell);
        if (!value && styleId === 0) return '';
        const r = `${this.columnToLetters(cIdx)}${rIdx + 1}`;
        const typeAttr = type ? ` t="${type}"` : '';
        const styleAttr = styleId > 0 ? ` s="${styleId}"` : '';
        return `<c r="${r}"${typeAttr}${styleAttr}><v>${value}</v></c>`;
      }).join('');
      return cells ? `<row r="${rIdx + 1}">${cells}</row>` : '';
    }).join('');

    const merges = sheet.merges && sheet.merges.length > 0
      ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map(ref => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>`
      : '';

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${cols}<sheetData>${rows}</sheetData>${merges}</worksheet>`;
  }

  private generateStyles(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${this.numFmts.length ? `<numFmts count="${this.numFmts.length}">${this.numFmts.join('')}</numFmts>` : ''}
  <fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>
  <fills count="${this.fills.length}">${this.fills.join('')}</fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="${this.cellXfs.length}">${this.cellXfs.join('')}</cellXfs>
</styleSheet>`;
  }

  private generateSharedStrings(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${this.sharedStrings.length}" uniqueCount="${this.sharedStrings.length}">${this.sharedStrings.map(s => `<si><t>${this.escapeXml(s)}</t></si>`).join('')}</sst>`;
  }

  private escapeXml(str: string): string {
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  }

  generate(): Uint8Array {
    const zip = new MiniZip();
    const worksheetXml = this.worksheets.map(sheet => this.generateWorksheet(sheet));
    zip.addFile('[Content_Types].xml', this.generateContentTypes());
    zip.addFile('_rels/.rels', this.generateRels());
    zip.addFile('xl/workbook.xml', this.generateWorkbook());
    zip.addFile('xl/_rels/workbook.xml.rels', this.generateWorkbookRels());
    zip.addFile('xl/styles.xml', this.generateStyles());
    zip.addFile('xl/sharedStrings.xml', this.generateSharedStrings());
    worksheetXml.forEach((xml, i) => zip.addFile(`xl/worksheets/sheet${i + 1}.xml`, xml));
    return zip.generate();
  }
}
