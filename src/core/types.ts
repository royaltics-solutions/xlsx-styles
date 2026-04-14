export interface Cell {
  value: string | number | boolean | Date | null;
  style?: CellStyle;
}

export interface CellStyle {
  font?: {
    bold?: boolean;
    italic?: boolean;
    size?: number;
    color?: string;
  };
  fill?: {
    color?: string;
  };
  alignment?: {
    horizontal?: 'left' | 'center' | 'right';
    vertical?: 'top' | 'middle' | 'bottom';
  };
  numberFormat?: string;
}

export interface Worksheet {
  name: string;
  data: (Cell | string | number | boolean | null)[][];
  columnWidths?: number[];
  merges?: string[];
}
