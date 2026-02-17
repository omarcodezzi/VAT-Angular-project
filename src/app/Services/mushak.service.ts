import { Injectable } from '@angular/core';
import { FinalRow } from './types'; // Ensure this path is correct

export interface Mushak91Data {
  note4: { value: number; sd: number; vat: number };
  note14: { value: number; vat: number };
  totalPayableVat: number;
}

@Injectable({ providedIn: 'root' })
export class MushakService {
  calculateNotes(rows: FinalRow[]): Mushak91Data {
    // Note 4: Standard Rated Goods (Section 3) [cite: 31]
    const note4 = {
      value: rows.reduce((s, r) => s + Number(r.VAT > 0 ? (r.CD || 0) : 0), 0), // Adjust logic if you use a specific 'Value' field
      sd: rows.reduce((s, r) => s + Number(r.SD || 0), 0),
      vat: rows.reduce((s, r) => s + Number(r.VAT || 0), 0)
    };

    // Note 14: Standard Rated Local Purchase (Section 4) [cite: 36]
    const note14 = {
      value: 0, // Set static or map from specific purchase rows
      vat: 0
    };

    return {
      note4,
      note14,
      // Section-7: (9C - 23B + 28 - 33) [cite: 41]
      totalPayableVat: note4.vat - note14.vat 
    };
  }
}