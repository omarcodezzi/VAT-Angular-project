import { Component, computed, ElementRef, signal, ViewChild } from "@angular/core";
import { CommonModule } from "@angular/common";
import { bestSuggestion, normalizeHeader } from "../Services/string-utils";
import { ExcelService } from "../Services/excel.service";
import { FinalRow, HeaderCheckRow, REQUIRED_HEADERS, TaxHeader } from "../Services/types";
import * as XLSX from 'xlsx';
import { HttpClient } from "@angular/common/http";
import { ExportService } from "../Services/export.service";
import { MushakService } from "../Services/mushak.service";

@Component({
  selector: "app-excel-import",
  standalone: true,
  imports: [CommonModule],
  templateUrl: "./excel-import.component.html",
  styleUrl: "./excel-import.component.css",
})
export class ExcelImportComponent {
  private excel = new ExcelService();

  excelHeaders = signal<string[]>([]);
  excelRows = signal<any[]>([]);
  showFinalGrid = signal<boolean>(false);

  headerMap = signal<Partial<Record<TaxHeader, string>>>({
    HSCode: undefined,
    Description: undefined,
    CD: undefined,
    SD: undefined,
    RD: undefined,
    VAT: undefined,
    AIT: undefined,
    TTI: undefined
    // Value, PurchaseValue, etc. are now optional, so no error here!
  });

  constructor(private http: HttpClient, private exportService: ExportService, private mushakService: MushakService) {

  }

  headerChecks = computed<HeaderCheckRow[]>(() => {
    const headers = this.excelHeaders();
    const map = this.headerMap();

    return REQUIRED_HEADERS.map((expected) => {
      const expN = normalizeHeader(expected);

      const manualMapping = map[expected];
      const exactMatch = headers.find((h) => normalizeHeader(h) === expN);
      const isMatched = !!(manualMapping || exactMatch);
      const status = isMatched ? "OK" : "MISSING";

      const found = isMatched ? expected : (exactMatch || "-");

      return {
        expected,
        found,
        status,
        suggestion: bestSuggestion(expected, headers),
        message: status === "OK" ? "Matched" : "Column missing"
      };
    });
  });



  canShowImportButton = computed(() => {
    console.log("Header Checks Status:", this.headerChecks());

    return this.headerChecks().every((row) => row.status === "OK");
  });


  finalRows = computed<FinalRow[]>(() => {
    if (!this.showFinalGrid() || !this.canShowImportButton()) return [];

    const rows = this.excelRows();
    const map = this.headerMap(); // This contains your manual mapping

    return rows.map((r, idx) => {
      const getValue = (key: string) => {
        // 1. Find the correct Excel header (e.g., 'TTII' for 'TTI')
        // Use the map if it exists, otherwise fallback to the key itself
        const excelHeader = map[key as TaxHeader] || key;
        const val = r[excelHeader];

        // 2. Strict check to allow 0 and 10 to show up
        return (val !== undefined && val !== null && val !== '') ? val : 0;
      };

      return {
        __rowIndex: idx + 2,
        HSCode: getValue('HSCode'),
        Description: getValue('Description'),
        CD: getValue('CD'),
        SD: getValue('SD'),
        VAT: getValue('VAT'),
        RD: getValue('RD'),
        AIT: getValue('AIT'),
        TTI: getValue('TTI') // This will now correctly pull from 'TTII'
      } as FinalRow;
    });
  });




  errorText = signal<string | null>(null);
  successText = signal<string | null>(null);

  onFileChange(event: any) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e: any) => {
      const binaryString = e.target.result;
      const workbook = XLSX.read(binaryString, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      const jsonData = XLSX.utils.sheet_to_json(worksheet);
      this.excelRows.set(jsonData);

      const headers = XLSX.utils.sheet_to_json(worksheet, { header: 1 })[0] as string[];
      this.excelHeaders.set(headers);

      this.showFinalGrid.set(false);
      this.updateHeaderMap();
    };
    reader.readAsBinaryString(file);
  }

  updateHeaderMap() {
    const headerCheckResults = this.headerChecks();

    headerCheckResults.forEach((row) => {
      if (row.status === "OK" && row.found) {
        this.headerMap.update((m) => ({ ...m, [row.expected]: row.found }));
      }
    });
  }


  applySuggestion(expected: TaxHeader, suggestion?: string) {
    if (!suggestion) return;

    // Save the SUGGESTION (e.g., 'AITI') to the map
    this.headerMap.update((m) => ({ ...m, [expected]: suggestion }));

    this.errorText.set(null);
    this.successText.set(`Suggestion applied for ${expected} ✅`);
  }

  importData() {
    debugger
    if (this.canShowImportButton()) {
      this.showFinalGrid.set(true);
      this.successText.set(`Ready to import ${this.finalRows().length} rows ✅`);
    }
  }


  onSelectSuggestion(expected: TaxHeader, event: any) {
    const selectedValue = event.target.value;
    if (!selectedValue) return;

    // Find the check row to get the 'suggestion' (the actual Excel column name)
    const checkRow = this.headerChecks().find(c => c.expected === expected);
    const actualExcelHeader = checkRow?.suggestion; // This will be 'AITI' or 'TTII'

    if (actualExcelHeader) {
      // Save the ACTUAL Excel header to the map
      this.headerMap.update((m) => ({ ...m, [expected]: actualExcelHeader }));
      this.successText.set(`Manual mapping applied for ${expected} ✅`);
      this.errorText.set(null);
    }
  }

  refresh(fileInput: HTMLInputElement) {
    this.errorText.set(null);
    this.successText.set(null);

    this.excelHeaders.set([]);
    this.excelRows.set([]);

    this.headerMap.set({
      HSCode: undefined,
      Description: undefined,
      CD: undefined,
      SD: undefined,
      RD: undefined,
      VAT: undefined,
      AIT: undefined,
      TTI: undefined,
    });
    this.showFinalGrid.set(false);
    fileInput.value = "";
  }

  saveFinalData() {
    const payload = this.finalRows().map(({ __rowIndex, ...rest }) => rest);

    this.http.post('/api/hs-code/import', payload).subscribe({
      next: () => this.successText.set(`Saved ${payload.length} rows ✅`),
      error: (err) => this.errorText.set(err?.message ?? 'Save failed'),
    });
  }

  downloadExcel() {
    const rows = this.finalRows();
    if (!rows.length) return;
    this.exportService.exportExcel(rows, 'VAT_HS_Code.xlsx');
  }

  downloadPdf() {
    const rows = this.finalRows();
    if (!rows.length) return;
    this.exportService.exportPdf(rows, 'VAT_HS_Code.pdf');
  }

downloadFullMushakPdf() {
  const data = this.exportService.mushakStaticData;
  if (!data || !data.notes) {
    console.error("Mushak static data is missing!");
    return;
  }
  this.exportService.exportFullMushakPdf(data);
}

  // For the Full Formatted Excel Report
  downloadFullMushakExcel() {
    this.exportService.exportFullMushakExcel(this.exportService.mushakStaticData);
  }
 
}
