export type TaxHeader = "HSCode" | "Description" | "CD" | "SD" | "RD" | "VAT" | "AIT" | "TTI" | "Value" | "PurchaseValue" | "PurchaseVat";

export const REQUIRED_HEADERS: TaxHeader[] = ["HSCode", "Description", "CD", "SD", "RD", "VAT", "AIT", "TTI"];

export interface ParsedSheet {
  headers: string[];
  rows: any[];
}
// Result per expected column
export interface HeaderCheckRow {
  expected: TaxHeader;
  found?: string;          // actual header found in Excel (if matched)
  status: "OK" | "MISSING";
  suggestion?: string;     // best suggestion from Excel headers
  message?: string;        // user-facing message
}

// Row of data after mapping headers → required keys
export type FinalRow = Record<TaxHeader, any> & { __rowIndex: number };
