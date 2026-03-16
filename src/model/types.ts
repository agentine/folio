export interface Color {
  argb?: string;
  theme?: number;
  tint?: number;
}

export interface Font {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
  strike?: boolean;
  color?: Color;
  family?: number;
  vertAlign?: 'superscript' | 'subscript';
}

export interface PatternFill {
  type: 'pattern';
  pattern: string;
  fgColor?: Color;
  bgColor?: Color;
}

export interface GradientFill {
  type: 'gradient';
  gradient: 'angle' | 'path';
  degree?: number;
  center?: { left: number; top: number };
  stops: { position: number; color: Color }[];
}

export type Fill = PatternFill | GradientFill;

export interface BorderStyle {
  style?: 'thin' | 'medium' | 'thick' | 'dotted' | 'dashed' | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot' | 'double' | 'hair';
  color?: Color;
}

export interface Border {
  top?: BorderStyle;
  right?: BorderStyle;
  bottom?: BorderStyle;
  left?: BorderStyle;
  diagonal?: BorderStyle & { up?: boolean; down?: boolean };
}

export interface Alignment {
  horizontal?: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
  vertical?: 'top' | 'middle' | 'bottom' | 'distributed' | 'justify';
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  textRotation?: number;
  readingOrder?: 'ltr' | 'rtl';
}

export interface Style {
  font?: Font;
  fill?: Fill;
  border?: Border;
  alignment?: Alignment;
  numFmt?: string;
  protection?: { locked?: boolean; hidden?: boolean };
}

export interface FormulaValue {
  formula: string;
  result?: string | number | boolean;
}

export interface RichTextRun {
  font?: Font;
  text: string;
}

export interface RichTextValue {
  richText: RichTextRun[];
}

export interface HyperlinkValue {
  text: string;
  hyperlink: string;
  tooltip?: string;
}

export interface ErrorValue {
  error: '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A';
}

export type CellValue =
  | null
  | string
  | number
  | boolean
  | Date
  | FormulaValue
  | RichTextValue
  | HyperlinkValue
  | ErrorValue;

export enum ValueType {
  Null = 0,
  String = 1,
  Number = 2,
  Date = 3,
  Boolean = 4,
  Formula = 5,
  RichText = 6,
  Hyperlink = 7,
  Error = 8,
}

export interface DataValidation {
  type?: 'list' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'custom';
  operator?: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';
  allowBlank?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  formulae?: string[];
  promptTitle?: string;
  prompt?: string;
  errorTitle?: string;
  error?: string;
  errorStyle?: 'stop' | 'warning' | 'information';
}

export interface SheetView {
  state?: 'frozen' | 'split' | 'normal';
  xSplit?: number;
  ySplit?: number;
  topLeftCell?: string;
  activeCell?: string;
  zoomScale?: number;
  showGridLines?: boolean;
  showRowColHeaders?: boolean;
  rightToLeft?: boolean;
}

export interface PageSetup {
  orientation?: 'portrait' | 'landscape';
  paperSize?: number;
  fitToPage?: boolean;
  fitToWidth?: number;
  fitToHeight?: number;
  margins?: {
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
    header?: number;
    footer?: number;
  };
  printArea?: string;
  printTitlesRow?: string;
  printTitlesColumn?: string;
  horizontalCentered?: boolean;
  verticalCentered?: boolean;
}

export interface ColumnDefinition {
  header?: string;
  key?: string;
  width?: number;
  style?: Style;
  hidden?: boolean;
  outlineLevel?: number;
}

export type SheetState = 'visible' | 'hidden' | 'veryHidden';
