import { ChangeDetectorRef, Component, ViewChild } from "@angular/core";
import {
  KENDO_SPREADSHEET,
  SheetDescriptor,
  SpreadsheetComponent,
} from "@progress/kendo-angular-spreadsheet";
import { sheets } from "./sheets";
import {
  exportSpreadsheetToPdf,
  SpreadsheetDocument,
} from "./spreadsheet-pdf.service";

@Component({
  selector: "my-app",
  imports: [KENDO_SPREADSHEET],
  template: `
    <div class="toolbar">
      <button type="button" class="k-button k-button-solid-primary" (click)="saveAsPdf()">
        Save as PDF
      </button>
    </div>
    <kendo-spreadsheet
      #spreadsheet
      [sheets]="sheets"
      (excelImport)="onExcelImport($event)"
      style="height: calc(100% - 48px); width: 100%"
    >
    </kendo-spreadsheet>
  `,
  styles: [
    `
      .toolbar {
        padding: 8px 12px;
        border-bottom: 1px solid var(--kendo-color-border, #e0e0e0);
        background: var(--kendo-color-app-surface, #fff);
      }
    `,
  ],
})
export class AppComponent {
  @ViewChild("spreadsheet") spreadsheetRef!: SpreadsheetComponent;
  public sheets: SheetDescriptor[] = [...sheets];

  constructor(private cdr: ChangeDetectorRef) {}

  onExcelImport(event: { file: File | Blob; preventDefault: () => void; sender: { fromFile: (f: File | Blob) => Promise<void>; toJSON: () => SpreadsheetDocument } }): void {
    event.preventDefault();
    event.sender.fromFile(event.file).then(() => {
      const doc = event.sender.toJSON();
      this.sheets = doc.sheets ?? [];
      this.cdr.detectChanges();
    });
  }

  saveAsPdf(): void {
    const widget = this.spreadsheetRef?.spreadsheetWidget;
    if (!widget) return;
    const doc: SpreadsheetDocument = widget.toJSON();
    const pdf = exportSpreadsheetToPdf(doc);
    pdf.save("Spreadsheet.pdf");
  }
}
