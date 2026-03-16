import { Component } from "@angular/core";
import {
  KENDO_SPREADSHEET,
  SheetDescriptor,
} from "@progress/kendo-angular-spreadsheet";
import { sheets } from "./sheets";
import { PdfExportService } from "./pdf-export.service";

@Component({
  selector: "my-app",
  imports: [KENDO_SPREADSHEET],
  template: `
    <div class="spreadsheet-wrapper">
      <div class="toolbar">
        <button type="button" class="save-pdf-btn" (click)="saveAsPdf()">
          Save as PDF
        </button>
      </div>
      <kendo-spreadsheet [sheets]="sheets" style="height: calc(100% - 48px); width: 100%">
      </kendo-spreadsheet>
    </div>
  `,
  styles: [
    `
      .spreadsheet-wrapper {
        height: 100%;
        width: 100%;
        display: flex;
        flex-direction: column;
      }
      .toolbar {
        flex: 0 0 auto;
        padding: 8px 12px;
        background: #f5f5f5;
        border-bottom: 1px solid #e0e0e0;
      }
      .save-pdf-btn {
        padding: 8px 16px;
        font-size: 14px;
        background: rgb(0, 62, 117);
        color: white;
        border: none;
        border-radius: 4px;
        cursor: pointer;
      }
      .save-pdf-btn:hover {
        background: rgb(0, 82, 147);
      }
    `,
  ],
})
export class AppComponent {
  public sheets: SheetDescriptor[] = sheets;

  constructor(private pdfExport: PdfExportService) {}

  saveAsPdf(): void {
    this.pdfExport.exportSheetToPdf(this.sheets, "invoice.pdf");
  }
}
