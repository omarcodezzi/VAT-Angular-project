import * as XLSX from "xlsx";

export interface ParsedSheet {
  headers: string[];
  rows: any[]; // array of objects by header
}

export class ExcelService {
  async readFirstSheet(file: File): Promise<ParsedSheet> {
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: "array" });

    const sheetName = wb.SheetNames[0];
    if (!sheetName) throw new Error("No sheet found in Excel.");

    const ws = wb.Sheets[sheetName];

    // 1) Read as arrays to get headers safely
    const raw: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (!raw.length) return { headers: [], rows: [] };

    const headers = (raw[0] ?? []).map((h) => (h ?? "").toString().trim());
    const body = raw.slice(1);

    // 2) Convert to objects using headers
    const rows = body
      .filter((r) => r.some((cell) => cell !== "")) // ignore empty rows
      .map((r) => {
        const obj: any = {};
        headers.forEach((h, i) => (obj[h] = r[i] ?? ""));
        return obj;
      });

    return { headers, rows };
  }
}
