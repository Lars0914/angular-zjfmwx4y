import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";

export interface SpreadsheetCell {
  value?: string | number | boolean | Date;
  formula?: string;
  index?: number;
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
}

const LETTER_LANDSCAPE_WIDTH_MM = 279.4;
const LETTER_LANDSCAPE_HEIGHT_MM = 215.9;
const MARGIN_LEFT = 10;
const MARGIN_RIGHT = 10;
const PRINTABLE_WIDTH_MM = LETTER_LANDSCAPE_WIDTH_MM - MARGIN_LEFT - MARGIN_RIGHT;
const TABLE_FONT_SIZE = 8;
const PX_TO_MM = 25.4 / 96;

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
    const w = widthByIndex.get(c) ?? PRINTABLE_WIDTH_MM / columnIndices.length;
    result.push(w);
    totalMm += w;
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

function sheetToMatrix(sheet: SpreadsheetSheet): string[][] {
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

  const matrix: string[][] = [];
  for (const rowIndex of sortedRowIndices) {
    const row = rowByIndex.get(rowIndex)!;
    const cells = row.cells ?? [];
    const cellByIndex = new Map<number, SpreadsheetCell>();
    for (const cell of cells) {
      const idx = cell.index ?? cellByIndex.size;
      cellByIndex.set(idx, cell);
    }
    const rowValues: string[] = [];
    for (let c = 0; c < maxCol; c++) {
      const cell = cellByIndex.get(c);
      const raw = cell?.value ?? cell?.formula;
      const val =
        raw === undefined || raw === null
          ? ""
          : typeof raw === "object" && (raw as Date).getTime
            ? (raw as Date).toISOString?.() ?? String(raw)
            : String(raw);
      rowValues.push(val);
    }
    matrix.push(rowValues);
  }
  return matrix;
}

function getColumnIndicesWithData(matrix: string[][]): number[] {
  const numCols = matrix[0]?.length ?? 0;
  const indices: number[] = [];
  for (let c = 0; c < numCols; c++) {
    for (let r = 0; r < matrix.length; r++) {
      const val = String(matrix[r]?.[c] ?? "").trim();
      if (val.length > 0) {
        indices.push(c);
        break;
      }
    }
  }
  return indices;
}

function removeEmptyColumns(
  matrix: string[][],
  columnIndices: number[]
): string[][] {
  if (columnIndices.length === 0) return [];
  return matrix.map((row) => columnIndices.map((c) => row[c] ?? ""));
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
    const matrix = removeEmptyColumns(rawMatrix, columnIndices);
    if (matrix.length === 0) continue;

    const [headRow, ...bodyRows] = matrix;
    const head = headRow ? [headRow] : [];
    const body = bodyRows;

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

  return pdf;
}
