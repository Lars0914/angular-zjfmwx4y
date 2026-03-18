import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";

export interface SpreadsheetCell {
  value?: string | number | boolean | Date;
  formula?: string;
  index?: number;
  format?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
}

export interface CellDisplay {
  value: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
}

export interface SpreadsheetRow {
  cells?: SpreadsheetCell[];
  index?: number;
}

export interface SpreadsheetColumn {
  index?: number;
  width?: number;
}

export interface SpreadsheetSheet {
  name?: string;
  rows?: SpreadsheetRow[];
  columns?: SpreadsheetColumn[];
}

export interface SpreadsheetDocument {
  sheets?: SpreadsheetSheet[];
  images?: { [id: string]: string };
}

const LETTER_LANDSCAPE_WIDTH_MM = 279.4;
const LETTER_LANDSCAPE_HEIGHT_MM = 215.9;
const MARGIN_LEFT = 10;
const MARGIN_RIGHT = 10;
const PRINTABLE_WIDTH_MM = LETTER_LANDSCAPE_WIDTH_MM - MARGIN_LEFT - MARGIN_RIGHT;
const TABLE_FONT_SIZE = 8;
const PX_TO_MM = 25.4 / 96;
const MAX_COLUMN_WIDTH_MM = 35;
const MIN_FIRST_COLUMN_WIDTH_MM = 9;

function formatNumberWithCommas(n: number, decimals = 2): string {
  const fixed = n.toFixed(decimals);
  const [intPart, decPart] = fixed.split(".");
  const withCommas = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return decPart != null ? `${withCommas}.${decPart}` : withCommas;
}

function isCurrencyFormat(format: string): boolean {
  const f = format.trim();
  return f.indexOf("$") !== -1 || f.indexOf("€") !== -1 || f.indexOf("¥") !== -1 || /\[$$\]|\[EUR\]|\[USD\]/i.test(f);
}

function isNumberWithCommasFormat(format: string): boolean {
  const f = format.trim();
  return f.indexOf(",") !== -1 && (f.indexOf("#") !== -1 || f.indexOf("0") !== -1);
}

function getDecimalsFromFormat(format: string): number {
  const f = format.trim();
  const decMatch = f.match(/\.([0#]+)/);
  if (!decMatch) return 0;
  return decMatch[1].length;
}

function getColumnWidthsMm(
  sheet: SpreadsheetSheet,
  columnIndices: number[]
): number[] {
  const columns = sheet.columns ?? [];
  if (columnIndices.length === 0) return [];

  const widthByIndex = new Map<number, number>();
  for (const col of columns) {
    const idx = col.index ?? widthByIndex.size;
    if (col.width != null && col.width > 0) {
      widthByIndex.set(idx, col.width * PX_TO_MM);
    }
  }

  if (widthByIndex.size === 0) {
    const w = PRINTABLE_WIDTH_MM / columnIndices.length;
    return new Array(columnIndices.length).fill(w);
  }

  const result: number[] = [];
  let totalMm = 0;
  for (const c of columnIndices) {
    let w = widthByIndex.get(c) ?? PRINTABLE_WIDTH_MM / columnIndices.length;
    w = Math.min(w, MAX_COLUMN_WIDTH_MM);
    result.push(w);
    totalMm += w;
  }
  if (result.length > 0) {
    const prevFirst = result[0];
    result[0] = Math.max(result[0], MIN_FIRST_COLUMN_WIDTH_MM);
    totalMm += result[0] - prevFirst;
  }

  if (totalMm <= 0) {
    const w = PRINTABLE_WIDTH_MM / columnIndices.length;
    return new Array(columnIndices.length).fill(w);
  }

  if (totalMm > PRINTABLE_WIDTH_MM) {
    const scale = PRINTABLE_WIDTH_MM / totalMm;
    return result.map((w) => w * scale);
  }

  if (totalMm < PRINTABLE_WIDTH_MM) {
    const remaining = PRINTABLE_WIDTH_MM - totalMm;
    return result.map((w) => w + (w / totalMm) * remaining);
  }

  return result;
}

function sheetToMatrix(sheet: SpreadsheetSheet): CellDisplay[][] {
  const rows = sheet.rows ?? [];
  if (rows.length === 0) return [];

  let maxCol = 0;
  for (const row of rows) {
    for (const cell of row.cells ?? []) {
      const idx = cell.index ?? 0;
      if (idx >= maxCol) maxCol = idx + 1;
    }
  }

  const rowByIndex = new Map<number, SpreadsheetRow>();
  for (const row of rows) {
    const i = row.index ?? rowByIndex.size;
    rowByIndex.set(i, row);
  }
  const sortedRowIndices = Array.from(rowByIndex.keys()).sort((a, b) => a - b);

  const matrix: CellDisplay[][] = [];
  for (const rowIndex of sortedRowIndices) {
    const row = rowByIndex.get(rowIndex)!;
    const cells = row.cells ?? [];
    const cellByIndex = new Map<number, SpreadsheetCell>();
    for (const cell of cells) {
      const idx = cell.index ?? cellByIndex.size;
      cellByIndex.set(idx, cell);
    }
    const rowValues: CellDisplay[] = [];
    for (let c = 0; c < maxCol; c++) {
      const cell = cellByIndex.get(c);
      const raw = cell?.value ?? cell?.formula;
      const format = cell?.format ?? "";
      let val: string;
      if (raw === undefined || raw === null) {
        val = "";
      } else if (typeof raw === "number") {
        const n = Number(raw);
        if (format.indexOf("%") !== -1) {
          const pct = Math.abs(n) <= 1 ? n * 100 : n;
          val = pct.toFixed(2) + "%";
        } else if (isCurrencyFormat(format)) {
          val = "$ " + formatNumberWithCommas(n);
        } else if (isNumberWithCommasFormat(format)) {
          val = formatNumberWithCommas(n, getDecimalsFromFormat(format));
        } else {
          val = Number.isInteger(n) ? String(n) : n.toFixed(2);
        }
      } else if (typeof raw === "object" && (raw as Date).getTime) {
        val = (raw as Date).toISOString?.() ?? String(raw);
      } else {
        val = String(raw);
      }
      rowValues.push({
        value: val,
        bold: cell?.bold,
        italic: cell?.italic,
        underline: cell?.underline,
      });
    }
    matrix.push(rowValues);
  }
  return matrix;
}

function getColumnIndicesWithData(matrix: CellDisplay[][]): number[] {
  const numCols = matrix[0]?.length ?? 0;
  const indices: number[] = [];
  for (let c = 0; c < numCols; c++) {
    for (let r = 0; r < matrix.length; r++) {
      const val = String(matrix[r]?.[c]?.value ?? "").trim();
      if (val.length > 0) {
        indices.push(c);
        break;
      }
    }
  }
  return indices;
}

function removeEmptyColumns(
  matrix: CellDisplay[][],
  columnIndices: number[]
): CellDisplay[][] {
  if (columnIndices.length === 0) return [];
  const empty: CellDisplay = { value: "" };
  return matrix.map((row) => columnIndices.map((c) => row[c] ?? empty));
}

function isRowEmpty(row: CellDisplay[]): boolean {
  return row.every((cell) => String(cell?.value ?? "").trim() === "");
}

function removeEmptyRows(matrix: CellDisplay[][]): CellDisplay[][] {
  if (matrix.length === 0) return [];
  const keep: boolean[] = matrix.map((row, i) => {
    const empty = isRowEmpty(row);
    if (!empty) return true;
    const nextRow = matrix[i + 1];
    return nextRow != null && !isRowEmpty(nextRow);
  });
  return matrix.filter((_, i) => keep[i]);
}

function cellFontStyle(cell: CellDisplay): FontStyle {
  if (cell?.bold && cell?.italic) return "bolditalic";
  if (cell?.bold) return "bold";
  if (cell?.italic) return "italic";
  return "normal";
}

type FontStyle = "normal" | "bold" | "italic" | "bolditalic";
type TableCellInput = string | { content: string; colSpan?: number; styles?: { fontStyle?: FontStyle } };
type TableRowInput = TableCellInput[];

function hasValue(row: CellDisplay[], i: number): boolean {
  return String(row[i]?.value ?? "").trim().length > 0;
}

function isEmptyFrom(row: CellDisplay[], start: number, numCols: number): boolean {
  for (let j = start; j < numCols; j++) {
    if (hasValue(row, j)) return false;
  }
  return true;
}

function rowToTableInput(row: CellDisplay[], numCols: number): TableRowInput {
  if (isRowEmpty(row)) {
    return [{ content: "", colSpan: numCols }];
  }

  const onlyFirstColumnHasValue =
    hasValue(row, 0) && isEmptyFrom(row, 1, numCols);

  if (onlyFirstColumnHasValue) {
    const cell = row[0];
    return [{ content: String(cell?.value ?? "").trim(), colSpan: numCols, styles: { fontStyle: cellFontStyle(cell ?? { value: "" }) } }];
  }

  const onlyFirstTwoHaveValues =
    hasValue(row, 0) &&
    hasValue(row, 1) &&
    isEmptyFrom(row, 2, numCols) &&
    numCols - 2 >= 3;
  const secondColValue = String(row[1]?.value ?? "").trim();
  const onlySecondColumnHasValueLong =
    numCols >= 2 &&
    !hasValue(row, 0) &&
    hasValue(row, 1) &&
    isEmptyFrom(row, 2, numCols) &&
    secondColValue.length > 60;

  if (onlyFirstTwoHaveValues || onlySecondColumnHasValueLong) {
    return [
      { content: String(row[0]?.value ?? "").trim(), colSpan: 1, styles: { fontStyle: cellFontStyle(row[0] ?? { value: "" }) } },
      { content: String(row[1]?.value ?? "").trim(), colSpan: numCols - 1, styles: { fontStyle: cellFontStyle(row[1] ?? { value: "" }) } },
    ];
  }
  return row.map((cell) => {
    const content = String(cell?.value ?? "").trim();
    const fontStyle = cellFontStyle(cell ?? { value: "" });
    return fontStyle !== "normal"
      ? { content, styles: { fontStyle } }
      : content;
  });
}

function rowUnderlines(row: CellDisplay[], numCols: number): boolean[] {
  if (isRowEmpty(row)) return [false];
  if (hasValue(row, 0) && isEmptyFrom(row, 1, numCols)) return [!!row[0]?.underline];
  if (
    (hasValue(row, 0) && hasValue(row, 1) && isEmptyFrom(row, 2, numCols) && numCols - 2 >= 3) ||
    (numCols >= 2 && !hasValue(row, 0) && hasValue(row, 1) && isEmptyFrom(row, 2, numCols) && String(row[1]?.value ?? "").trim().length > 60)
  ) {
    return [!!row[0]?.underline, !!row[1]?.underline];
  }
  return row.map((c) => !!c?.underline);
}

export function exportSpreadsheetToPdf(doc: SpreadsheetDocument): jsPDF {
  const pdf = new jsPDF({
    orientation: "landscape",
    unit: "mm",
    format: "letter",
  });

  const sheets = doc.sheets ?? [];
  let firstTable = true;

  for (let s = 0; s < sheets.length; s++) {
    const sheet = sheets[s];
    const rawMatrix = sheetToMatrix(sheet);
    const columnIndices = getColumnIndicesWithData(rawMatrix);
    let matrix = removeEmptyColumns(rawMatrix, columnIndices);
    matrix = removeEmptyRows(matrix);
    if (matrix.length === 0) continue;

    const numCols = matrix[0]?.length ?? 0;
    const [headRow, ...bodyRows] = matrix;
    const head = headRow
      ? [rowToTableInput(headRow, numCols)]
      : [];
    const body = bodyRows.map((row) => rowToTableInput(row, numCols));
    const headUnderlines = headRow ? [rowUnderlines(headRow, numCols)] : [];
    const bodyUnderlines = bodyRows.map((row) => rowUnderlines(row, numCols));
    const underlinesMatrix: boolean[][] = [...headUnderlines, ...bodyUnderlines];

    const columnWidths = getColumnWidthsMm(sheet, columnIndices);
    const tableTotalWidth = columnWidths.reduce((a, b) => a + b, 0);
    const columnStyles: Record<number, { cellWidth: number }> = {};
    columnWidths.forEach((w, i) => {
      columnStyles[i] = { cellWidth: w };
    });

    const opts: Parameters<typeof autoTable>[1] = {
      head,
      body,
      theme: "grid",
      margin: { left: MARGIN_LEFT, right: MARGIN_RIGHT },
      pageBreak: "auto",
      rowPageBreak: "avoid",
      showHead: head.length ? "everyPage" : "never",
      tableWidth: tableTotalWidth > 0 ? tableTotalWidth : PRINTABLE_WIDTH_MM,
      columnStyles,
      styles: { fontSize: TABLE_FONT_SIZE, cellPadding: 2 },
      headStyles: { fillColor: [240, 240, 240], fontSize: TABLE_FONT_SIZE },
      didDrawCell: (data: { doc: { getDocument: () => jsPDF }; row: { index: number }; column: { index: number }; cell: { x: number; y: number; width: number; height: number } }) => {
        const underline = underlinesMatrix[data.row.index]?.[data.column.index];
        if (underline && data.cell) {
          const doc = data.doc.getDocument();
          doc.setDrawColor(0, 0, 0);
          doc.setLineWidth(0.08);
          const y = data.cell.y + data.cell.height - 0.3;
          doc.line(data.cell.x, y, data.cell.x + data.cell.width, y);
        }
      },
    };

    let startY = 15;
    if (!firstTable) {
      pdf.addPage([LETTER_LANDSCAPE_WIDTH_MM, LETTER_LANDSCAPE_HEIGHT_MM], "landscape");
    } else {
      firstTable = false;
    }

    if (sheet.name) {
      pdf.setFontSize(12);
      pdf.text(sheet.name, 10, startY);
      startY += 8;
    }
    opts.startY = startY;

    autoTable(pdf, opts);
  }

  const imageIds = doc.images && Object.keys(doc.images);
  if (imageIds?.length) {
    const imgMargin = 10;
    const maxImgW = LETTER_LANDSCAPE_WIDTH_MM - 2 * imgMargin;
    const maxImgH = 100;
    const imgGap = 10;
    let y = 20;
    let needNewPage = true;
    for (const id of imageIds) {
      const src = doc.images![id];
      if (!src || typeof src !== "string") continue;
      if (!src.startsWith("data:")) continue;
      try {
        if (needNewPage) {
          pdf.addPage([LETTER_LANDSCAPE_WIDTH_MM, LETTER_LANDSCAPE_HEIGHT_MM], "landscape");
          y = 20;
          needNewPage = false;
        }
        const format = src.indexOf("image/jpeg") !== -1 || src.indexOf("image/jpg") !== -1 ? "JPEG" : "PNG";
        pdf.addImage(src, format, imgMargin, y, maxImgW, maxImgH);
        y += maxImgH + imgGap;
        if (y + maxImgH > LETTER_LANDSCAPE_HEIGHT_MM - 20) needNewPage = true;
      } catch {
        /* skip invalid image */
      }
    }
  }

  return pdf;
}
