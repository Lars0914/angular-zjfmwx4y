import { Component } from "@angular/core";
import {
  KENDO_SPREADSHEET,
  SheetDescriptor,
} from "@progress/kendo-angular-spreadsheet";
import { sheets } from "./sheets";

@Component({
  selector: "my-app",
  imports: [KENDO_SPREADSHEET],
  template: `
    <kendo-spreadsheet [sheets]="sheets" style="height: 100%; width: 100%">
    </kendo-spreadsheet>
  `,
})
export class AppComponent {
  public sheets: SheetDescriptor[] = sheets;
}
