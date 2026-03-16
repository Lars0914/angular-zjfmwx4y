import { Injectable } from "@angular/core";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import type { SheetDescriptor } from "@progress/kendo-angular-spreadsheet";

type CellLike = {
  value?: string | number;
  formula?: string;
  format?: string;
  background?: string;
  color?: string;
  textAlign?: string;
  bold?: boolean;
  fontSize?: number;
};
type RowLike = { cells?: CellLike[]; height?: number; index?: number };

/** Converts rgb(r,g,b) to [r,g,b] for jsPDF */
function parseRgb(str: string): [number, number, number] | null {
  if (!str || typeof str !== "string") return null;
  const m = str.match(/rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/);
  return m ? [parseInt(m[1], 10), parseInt(m[2], 10), parseInt(m[3], 10)] : null;
}

/** Format number as currency */
function formatCurrency(n: number): string {
  return "$" + n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/** Evaluate simple formula for one cell. rowValues = resolved values for current row [id, product, qty, price, tax, amount]. */
function evalFormula(
  formula: string,
  rowValues: (string | number)[]
): number | null {
  if (!formula) return null;
  const f = formula.toUpperCase();
  // C*D*0.2 -> tax
  const qty = typeof rowValues[2] === "number" ? rowValues[2] : parseFloat(String(rowValues[2] || "0"));
  const price = typeof rowValues[3] === "number" ? rowValues[3] : parseFloat(String(rowValues[3] || "0"));
  const tax = typeof rowValues[4] === "number" ? rowValues[4] : parseFloat(String(rowValues[4] || "0"));
  if (f.includes("C") && f.includes("D") && f.includes("0.2")) return qty * price * 0.2;
  if (f.includes("C") && f.includes("D") && f.includes("+E")) return qty * price + tax;
  return null;
}

/** Get display value for a cell. rowValues = resolved values for current row so far (for formula eval). */
function cellDisplayValue(cell: CellLike | Record<string, unknown>, rowValues: (string | number)[], isCurrency: boolean): string {
  const c = cell as CellLike;
  if (c.value !== undefined && c.value !== null) {
    const v = c.value;
    if (typeof v === "number" && (isCurrency || c.format?.includes("$"))) return formatCurrency(v);
    return String(v);
  }
  if (c.formula) {
    const n = evalFormula(c.formula, rowValues);
    if (n !== null && (isCurrency || c.format?.includes("$"))) return formatCurrency(n);
    if (n !== null) return String(n);
  }
  return "";
}

@Injectable({ providedIn: "root" })
export class PdfExportService {
  /**
   * Build and download a nice invoice-style PDF from the first sheet of the given workbook.
   */
  exportSheetToPdf(sheets: SheetDescriptor[], filename = "invoice.pdf"): void {
    const sheet = sheets?.[0];
    if (!sheet?.rows?.length) {
      console.warn("No sheet data for PDF export");
      return;
    }

    const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
    const pageW = doc.internal.pageSize.getWidth();
    const margin = 14;
    let y = margin;

    // 1) Title row (merged A1:G1 style - first row, first cell)
    const titleRow = sheet.rows[0];
    const titleCell = titleRow?.cells?.[0];
    const title =
      titleCell?.value != null ? String(titleCell.value) : sheet.name || "Invoice";
    doc.setFillColor(96, 181, 255);
    doc.rect(0, 0, pageW, 24, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(18);
    doc.setFont("helvetica", "bold");
    doc.text(title, pageW / 2, 14, { align: "center" });
    y = 28;

    // 2) Build table from rows (skip title row; treat row 1 as header, then data, then total)
    const headerRow = sheet.rows[1];
    const dataRows = sheet.rows.slice(2);
    const headerCells = headerRow?.cells ?? [];
    const headers = headerCells
      .map((c) => (c.value != null ? String(c.value) : ""))
      .filter(Boolean);
    const numCols = headers.length;

    // Line-item rows only (exclude Tip and Total rows for body)
    const lineItemCount = Math.max(0, dataRows.length - 2);
    const lineItemRows = dataRows.slice(0, lineItemCount);

    const bodyRows: (string | number)[][] = [];
    let totalAmount = 0;
    for (let r = 0; r < lineItemRows.length; r++) {
      const row = lineItemRows[r];
      const cells = row?.cells ?? [];
      const values: (string | number)[] = [];
      for (let col = 0; col < numCols; col++) {
        const cell = cells[col] ?? cells.find((c) => (c as { index?: number }).index === col);
        const isCurrency = col >= 3;
        const display = cellDisplayValue((cell ?? {}) as CellLike, values, isCurrency);
        const num = display.replace(/[$,]/g, "");
        const parsed = isCurrency && !Number.isNaN(parseFloat(num)) ? parseFloat(num) : display;
        values.push(parsed);
        if (col === 5 && typeof parsed === "number") totalAmount += parsed;
      }
      bodyRows.push(values.map((v) => String(v)));
    }

    // Total row
    const totalRow = headers.map((_, i) =>
      i === 2 ? "Total Amount:" : i === 5 ? formatCurrency(totalAmount) : ""
    );
    const body = [...bodyRows, totalRow];

    autoTable(doc, {
      startY: y,
      head: [headers],
      body,
      theme: "plain",
      margin: { left: margin, right: margin },
      headStyles: {
        fillColor: [167, 214, 255],
        textColor: [0, 62, 117],
        fontStyle: "bold",
        halign: "center",
      },
      bodyStyles: {
        textColor: [0, 62, 117],
        fontSize: 10,
      },
      alternateRowStyles: {
        fillColor: [229, 243, 255],
      },
      columnStyles: {
        0: { halign: "center" },
        2: { halign: "center" },
        3: { halign: "right" },
        4: { halign: "right" },
        5: { halign: "right" },
      },
      didParseCell: (data) => {
        if (data.section === "body" && data.row.index === body.length - 1) {
          data.cell.styles.fillColor = [193, 226, 255];
          data.cell.styles.fontStyle = "bold";
          data.cell.styles.fontSize = 11;
        }
      },
    });

    doc.save(filename);
  }
}
