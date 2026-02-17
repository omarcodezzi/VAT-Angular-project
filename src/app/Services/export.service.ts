import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

// pdfmake imports
import pdfMakeImport from 'pdfmake/build/pdfmake';
import pdfFontsImport from 'pdfmake/build/vfs_fonts';

// Make pdfMake and pdfFonts mutable
const pdfMake: any = pdfMakeImport;
const pdfFonts: any = pdfFontsImport;

// Resolve VFS fonts correctly
const vfs =
  pdfFonts?.pdfMake?.vfs ??
  pdfFonts?.vfs ??
  pdfFonts?.default?.pdfMake?.vfs ??
  pdfFonts?.default?.vfs;

// Set VFS in pdfMake if found
if (!vfs) {
  console.warn('pdfmake VFS not found. Check pdfmake/vfs_fonts import.');
} else {
  pdfMake.vfs = vfs;
}

@Injectable({
  providedIn: 'root',
})
export class ExportService {

  // STATIC DATA matching the Mushak-9.1 Form
  public mushakStaticData = {
    taxpayer: {
      bin: '123456789',
      name: 'VAT Learning Mission (DTCL)',
      address: 'Dhaka, Bangladesh',
      businessNature: 'Proprietorship',
      activity: 'Retail/Wholesale, Trading'
    },
    returnSubmission: {
      period: 'Oct / 2022',
      type: 'A) Main/Original Return (Section 64)',
      date: '03-Oct-2022'
    },
    notes: {
      note4: { val: 159270.30, sd: 0.00, vat: 23890.55 },
      note9: { val: 159270.30, sd: 0.00, vat: 23890.55 },
      note14: { val: 3717678.34, vat: 557651.75 },
      note23: { val: 3717678.34, vat: 557651.75 },
      note28: 0.00, note33: 0.00, note34: -533761.20,
      note35: 1979177.91, note50: 1979177.91, note52: 1445416.71,
      note54: 0.00, note65: 1979177.91, note67: 0.00, note68: 0.00
    }
  };

  // --- 1. HS CODE EXPORTS ---

  async exportExcel(rows: any[], fileName = 'VAT_HS_Code.xlsx') {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('HS Codes');

    sheet.columns = [
      { header: 'HS Code', key: 'HSCode', width: 15 },
      { header: 'Description', key: 'Description', width: 45 },
      { header: 'CD', key: 'CD', width: 8 },
      { header: 'SD', key: 'SD', width: 8 },
      { header: 'VAT', key: 'VAT', width: 8 },
      { header: 'AIT', key: 'AIT', width: 8 },
      { header: 'RD', key: 'RD', width: 8 },
      { header: 'TTI', key: 'TTI', width: 8 },
    ];

    rows.forEach((r) => {
      sheet.addRow({
        HSCode: r.HSCode ?? '',
        Description: r.Description ?? '',
        CD: r.CD ?? 0,
        SD: r.SD ?? 0,
        VAT: r.VAT ?? 0,
        AIT: r.AIT ?? 0,
        RD: r.RD ?? 0,
        TTI: r.TTI ?? 0,
      });
    });

    const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
    sheet.eachRow((row) => {
      row.eachCell((cell) => {
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' },
          bottom: { style: 'thin' }, right: { style: 'thin' },
        };
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), fileName);
  }

  exportPdf(rows: any[], fileName = 'VAT_HS_Code.pdf') {
    const body = [
      ['HS Code', 'Description', 'CD', 'SD', 'VAT', 'AIT', 'RD', 'TTI'],
      ...rows.map((r) => [
        String(r.HSCode ?? ''), String(r.Description ?? ''), String(r.CD ?? 0),
        String(r.SD ?? 0), String(r.VAT ?? 0), String(r.AIT ?? 0),
        String(r.RD ?? 0), String(r.TTI ?? 0),
      ]),
    ];

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [20, 30, 20, 30],
      content: [
        { text: 'VAT HS Code', style: 'title' },
        { table: { headerRows: 1, widths: [60, '*', 30, 30, 35, 35, 30, 30], body }, layout: 'lightHorizontalLines' },
      ],
      styles: { title: { fontSize: 14, bold: true, margin: [0, 0, 0, 8] } },
      defaultStyle: { fontSize: 9 },
    };
    (pdfMake as any).createPdf(docDef).download(fileName);
  }

  // --- 2. MUSHAK-9.1 EXPORTS (SECTIONS 1-8) ---
  private createFullWidthHeader(text: string) {
    return {
      table: {
        widths: ['*'],
        body: [[{ text: text, style: 'secHeaderCell' }]]
      },
      layout: 'noBorders',
      margin: [0, 2, 0, 0]
    };
  }

  exportFullMushakPdf(data: any) {
    const n = data?.notes || {};
    const t = data?.taxpayer || {};
    const s = data?.returnSubmission || {};

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      content: [
        { stack: [{ text: "GOVERNMENT OF THE PEOPLE'S REPUBLIC OF BANGLADESH", style: 'header' }, { text: "NATIONAL BOARD OF REVENUE", style: 'header' }, { text: "\nVALUE ADDED TAX RETURN FORM (Mushak-9.1)", style: 'subHeader' }] },

        this.createFullWidthHeader("SECTION - 1: TAXPAYER'S INFORMATION"),
        {
          style: 'dataTable',
          table: {
            widths: ['40%', '2%', '58%'],
            body: [
              ['1. BIN', ':', t.bin],
              ['2. Name', ':', t.name],
              ['3. Address', ':', t.address || ''],
              ['4. Nature', ':', t.businessNature],
              ['5. Activity', ':', t.activity]
            ]
          }
        },

        this.createFullWidthHeader("SECTION - 2: RETURN SUBMISSION DATA"),
        {
          style: 'dataTable',
          table: {
            widths: ['40%', '2%', '58%'],
            body: [
              ['1. Tax Period', ':', { text: s.period || 'Oct / 2022', alignment: 'center' }],
              ['2. Type of Return\n[Please select your desired option]', ':', {
                stack: [
                  { columns: [{ width: '*', text: 'A) Main/Original Return (Section 64)' }, { width: 40, text: '[   ]', alignment: 'right' }] },
                  { columns: [{ width: '*', text: 'B) Late Return (section 65)' }, { width: 40, text: '[   ]', alignment: 'right' }] },
                  { columns: [{ width: '*', text: 'C) Amend Return (section 66)' }, { width: 40, text: '[   ]', alignment: 'right' }] },
                  { columns: [{ width: '*', text: 'D) Full or Additional or Alternative Return (Section 67)' }, { width: 40, text: '[   ]', alignment: 'right' }] }
                ], margin: [0, 5, 0, 5]
              }],
              ['3. Any activities in this Tax Period?', ':', {
                stack: [
                  { columns: [{ width: 'auto', text: '[  ] Yes' }, { width: 30, text: '' }, { width: 'auto', text: '[   ] No' }], alignment: 'center' },
                  { text: '[If Selected "No" Please Fill Only Section I, II & X]', fontSize: 7, alignment: 'center', margin: [0, 2] }
                ]
              }],
              ['4. Date of Submission', ':', { text: s.date || '03-Oct-2022', alignment: 'center' }]
            ]
          }
        },

        this.createFullWidthHeader("SECTION - 3: SUPPLY - OUTPUT TAX"),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['35%', '10%', '5%', '15%', '13%', '13%', '9%'],
            body: [
              // Table Header
              [
                { text: 'Nature of Supply', style: 'tHead', colSpan: 2, alignment: 'center' },
                {},
                { text: 'Note', style: 'tHead', alignment: 'center' },
                { text: 'Value (a)', style: 'tHead', alignment: 'center' },
                { text: 'SD (b)', style: 'tHead', alignment: 'center' },
                { text: 'VAT (c)', style: 'tHead', alignment: 'center' },
                { text: '', style: 'tHead', border: [false, false, false, false] }
              ],
              // Row 1 & 2: Zero Rated
              [
                { text: 'Zero Rated Goods/Service', rowSpan: 2 },
                'Direct Export', '1', '0.00', { text: '', fillColor: '#d9d9d9' }, { text: '', fillColor: '#d9d9d9' }, 'Sub form'
              ],
              [
                {}, 'Deemed Export', '2', '0.00', { text: '', fillColor: '#d9d9d9' }, { text: '', fillColor: '#d9d9d9' }, 'Sub form'
              ],
              // Row 3: Exempted
              [{ text: 'Exempted Goods/Service', colSpan: 2 }, {}, '3', '0.00', { text: '', fillColor: '#d9d9d9' }, { text: '', fillColor: '#d9d9d9' }, 'Sub form'],
              // Row 4: Standard Rated
              [
                { text: 'Standard Rated Goods/Service', colSpan: 2 }, {}, '4',
                (n.note4?.val || 159270.30).toLocaleString(undefined, { minimumFractionDigits: 2 }),
                '0.00',
                (n.note4?.vat || 23890.55).toLocaleString(undefined, { minimumFractionDigits: 2 }),
                'Sub form'
              ],
              // Rows 5-7
              [{ text: 'Goods Based on MRP', colSpan: 2 }, {}, '5', '', '', '', 'Sub form'],
              [{ text: 'Goods/Service Based on Specific VAT', colSpan: 2 }, {}, '6', '', '', '', 'Sub form'],
              [{ text: 'Goods/Service Other than Standard Rate', colSpan: 2 }, {}, '7', '', '', '', 'Sub form'],
              // Row 8
              [{ text: 'Retail/Whole Sale/Trade Based Supply', colSpan: 2 }, {}, '8', '0.00', '0.00', '0.00', { text: 'Sub form', fillColor: '#d9d9d9' }],
              // Row 9: Total
              [
                { text: 'Total Sales Value & Total Payable Taxes', colSpan: 2, style: 'tBold' }, {},
                { text: '9', style: 'tBold' },
                { text: (n.note9?.val || 159270.30).toLocaleString(undefined, { minimumFractionDigits: 2 }), style: 'tBold', fillColor: '#d9d9d9' },
                { text: '0.00', style: 'tBold', fillColor: '#d9d9d9' },
                { text: (n.note9?.vat || 23890.55).toLocaleString(undefined, { minimumFractionDigits: 2 }), style: 'tBold', fillColor: '#d9d9d9' },
                { text: '', border: [false, false, false, false] }
              ]
            ]
          }
        },

        this.createFullWidthHeader("SECTION - 4: PURCHASE - INPUT TAX"),
        {
          stack: [
            {
              canvas: [{ type: 'rect', x: 0, y: 0, w: 535, h: 55, color: '#fcd5b4' }]
            },
            {
              text: [
                "1) If all the products/services you supply are standard rated, fill up note 10-20.\n",
                "2) All the products/services you supply are not standard rated or input tax credit not taken within stipulated time period under section 46, fill up note 21-22.\n",
                "3) If the products/services you supply consist of both standard rated and non-standard rated, then fill up note 10-20 for the raw materials that were used to produce/supply standard rated goods/services and fill up note 21-22 for the raw materials that were used to produce/supply non-standard rated goods/services and show the value proportionately in note 10-22 as applicable."
              ],
              fontSize: 7,
              margin: [5, -50, 5, 10]
            }
          ],
          // margin: [0, 5, 0, 10]
        },
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['35%', '15%', '5%', '17%', '17%', '11%'],
            body: [
              [
                { text: 'Nature of Purchase', style: 'tHead', colSpan: 2, alignment: 'center' },
                {},
                { text: 'Note', style: 'tHead', alignment: 'center' },
                { text: 'Value (a)', style: 'tHead', alignment: 'center' },
                { text: 'VAT (b)', style: 'tHead', alignment: 'center' },
                { text: '', border: [false, false, false, false] }
              ],
              // Zero Rated & Exempted (Notes 10-13)
              [{ text: 'Zero Rated Goods/Service', rowSpan: 2 }, 'Local Purchase', '10', '0.00', { text: '', fillColor: '#d9d9d9' }, 'Sub form'],
              [{}, 'Import', '11', '0.00', { text: '', fillColor: '#d9d9d9' }, 'Sub form'],
              [{ text: 'Exempted Goods/Service', rowSpan: 2 }, 'Local Purchase', '12', '0.00', { text: '', fillColor: '#d9d9d9' }, 'Sub form'],
              [{}, 'Import', '13', '0.00', { text: '', fillColor: '#d9d9d9' }, 'Sub form'],

              // Standard Rated - Main Data (Notes 14-15)
              [{ text: 'Standard Rated Goods/Service', rowSpan: 2 }, 'Local Purchase', '14', (n.note14?.val || 3717678.34).toLocaleString(), (n.note14?.vat || 557651.75).toLocaleString(), 'Sub form'],
              [{}, 'Import', '15', '0.00', '0.00', 'Sub form'],

              // Other Categories (Notes 16-22)
              [{ text: 'Goods/Service Other than Standard Rate', rowSpan: 2 }, 'Local Purchase', '16', '0.00', '0.00', 'Sub form'],
              [{}, 'Import', '17', '0.00', '0.00', 'Sub form'],
              [{ text: 'Goods/Service Based on Specific VAT', rowSpan: 1 }, 'Local Purchase', '18', '0.00', '0.00', 'Sub form'],
              [{ text: 'Goods/Service Not Admissible for Credit (Local Purchase)', rowSpan: 2 }, 'From Turnover Units', '19', '0.00', '0.00', 'Sub form'],
              [{}, 'From Unregistered Entities', '20', '0.00', '0.00', 'Sub form'],
              [{ text: 'Goods/Service Not Admissible for Credit (Taxpayers who sell only Exempted/ Specific VAT and Goods/Service Other than Standard Rate/\nCredits not taken\n', rowSpan: 2 }, 'Local Purchase', '21', '0.00', '0.00', 'Sub form'],
              [{}, 'Import', '22', '0.00', '0.00', 'Sub form'],

              // Total Row (Note 23)
              [
                { text: 'Total Input Tax Credit', colSpan: 1, style: 'tBold' },
                {},
                { text: '23', style: 'tBold' },
                { text: (n.note23?.val || 3717678.34).toLocaleString(), style: 'tBold', fillColor: '#d9d9d9' },
                { text: (n.note23?.vat || 557651.75).toLocaleString(), style: 'tBold', fillColor: '#d9d9d9' },
                { text: '', border: [false, false, false, false] }
              ]
            ]
          }
        },

        { text: '', pageBreak: 'before' },
        this.createFullWidthHeader("SECTION - 5: INCREASING ADJUSTMENTS"),
        { style: 'dataTable', table: { widths: ['*', 40, 100], body: [['Total Increasing Adjustment', '28', (n.note28 || 0).toLocaleString()]] } },

        this.createFullWidthHeader("SECTION - 6: DECREASING ADJUSTMENTS"),
        { style: 'dataTable', table: { widths: ['*', 40, 100], body: [['Total Decreasing Adjustment', '33', (n.note33 || 0).toLocaleString()]] } },

        this.createFullWidthHeader("SECTION - 7: NET TAX CALCULATION"),
        { style: 'dataTable', table: { widths: ['*', 40, 100], body: [['Net Payable VAT (34)', '34', (n.note34 || 0).toLocaleString()], ['For Treasury (50)', '50', (n.note50 || 0).toLocaleString()], ['Closing Balance (52)', '52', (n.note52 || 0).toLocaleString()]] } },

        this.createFullWidthHeader("SECTION - 8: OLD ACCOUNT BALANCE"),
        { style: 'dataTable', table: { widths: ['*', 40, 100], body: [['Balance from Mushak-18.6', '54', (n.note54 || 0).toLocaleString()]] } },

        { text: '', pageBreak: 'before' },
        this.createFullWidthHeader("SECTION - 9: ACCOUNT CODE WISE PAYMENT SCHEDULE"),
        { style: 'dataTable', table: { widths: ['*', 35, 120, 80, 45], body: [[{ text: 'Items', style: 'tHead' }, 'Note', 'Code', 'Amount', ''], ['VAT Deposit', '58', '1/1133/0030/0311', '0.00', 'Sub form']] } },

        this.createFullWidthHeader("SECTION - 10: CLOSING BALANCE"),
        { style: 'dataTable', table: { widths: ['*', 40, 100], body: [['Closing Balance (VAT)', '65', (n.note65 || 0).toLocaleString()], ['Closing Balance (SD)', '66', '0.00']] } },

        this.createFullWidthHeader("SECTION - 11: REFUND"),
        { style: 'dataTable', table: { widths: ['*', 40, 100], body: [['Requested Refund (VAT)', '67', (n.note67 || 0).toLocaleString()], ['Requested Refund (SD)', '68', (n.note68 || 0).toLocaleString()]] } },

        this.createFullWidthHeader("SECTION - 12: DECLARATION"),
        { text: "I hereby declare that all information provided in this Return Form are complete, true & accurate.", margin: [0, 10], fontSize: 8 },
        { table: { widths: ['25%', '5%', '70%'], body: [['Name', ':', 'Hasanuzzaman'], ['Signature', ':', '']] }, layout: 'noBorders' }
      ],
      styles: {
        header: { fontSize: 10, bold: true, alignment: 'center' },
        subHeader: { fontSize: 9, bold: true, alignment: 'center', color: '#003366' },
        secHeaderCell: { fillColor: '#003366', color: 'white', bold: true, alignment: 'center', fontSize: 9, padding: [0, 2, 0, 2] },
        tHead: { fillColor: '#f2f2f2', bold: true, fontSize: 8 },
        tBold: { bold: true, fontSize: 8 },
        dataTable: { fontSize: 8, margin: [0, 0, 0, 5] },
        borderedTable: { margin: [0, 0, 0, 2] }
      }
    };
    pdfMake.createPdf(docDef).download('Mushak_9.1_Full_Report.pdf');
  }

  // --- MERGED MUSHAK-9.1 EXCEL (ALL SECTIONS 1-12) ---
  async exportFullMushakExcel(data: any) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Mushak-9.1');

    // --- GOVERNMENT BRANDING HEADER ---
    const brandRow1 = sheet.addRow(["GOVERNMENT OF THE PEOPLE'S REPUBLIC OF BANGLADESH", '', '']);
    sheet.mergeCells(`B${brandRow1.number}:E${brandRow1.number}`);
    brandRow1.getCell(1).font = { size: 12, bold: true };
    brandRow1.getCell(1).alignment = { horizontal: 'center' };

    const brandRow2 = sheet.addRow(["NATIONAL BOARD OF REVENUE", '', '']);
    sheet.mergeCells(`B${brandRow2.number}:E${brandRow2.number}`);
    brandRow2.getCell(1).font = { size: 11, bold: true };
    brandRow2.getCell(1).alignment = { horizontal: 'center' };

    const formTitleRow = sheet.addRow(["VALUE ADDED TAX RETURN FORM", '', '|| Mushak-9.1 ||']);
    sheet.mergeCells(`B${formTitleRow.number}:E${formTitleRow.number}`);
    formTitleRow.getCell(1).font = { size: 10, bold: true };
    formTitleRow.getCell(1).alignment = { horizontal: 'center' };
    formTitleRow.getCell(6).font = { size: 10, bold: true }; // Mushak-9.1 ID on right

    const ruleRow = sheet.addRow(["[Rule 47(1)]", '', '']);
    sheet.mergeCells(`B${ruleRow.number}:E${ruleRow.number}`);
    ruleRow.getCell(1).font = { size: 8 };
    ruleRow.getCell(1).alignment = { horizontal: 'center' };

    // --- 1. COLUMN SETUP ---
    sheet.columns = [
      { width: 35 }, // A: Label
      { width: 3 },  // B: Separator (:)
      { width: 35 }, // C: Data
      { width: 10 }, // D: Spacing
      { width: 10 }, // E: Spacing
      { width: 12 }  // F: Sub form
    ];

    // --- 2. STYLING HELPERS ---
    const addHeader = (text: string) => {
      sheet.addRow([]); // Spacer
      const row = sheet.addRow([text]);
      sheet.mergeCells(`A${row.number}:E${row.number}`);
      row.eachCell(c => {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '003366' } };
        c.font = { color: { argb: 'FFFFFF' }, bold: true };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
      });
    };

    const applyBorder = (start: number, end: number) => {
      for (let i = start; i <= end; i++) {
        sheet.getRow(i).eachCell({ includeEmpty: true }, (c, col) => {
          if (col <= 5) {
            c.border = {
              top: { style: 'thin' }, left: { style: 'thin' },
              bottom: { style: 'thin' }, right: { style: 'thin' }
            };
          }
        });
      }
    };

    // --- SECTION 1 & 2 ---
    addHeader("SECTION - 1: TAXPAYER'S INFORMATION");
    const s1Start = sheet.rowCount + 1;
    const s1Data = [
      ['1. BIN', ':', data.taxpayer.bin],
      ['2. Name of Taxpayer', ':', data.taxpayer.name],
      ['3. Address of Taxpayer', ':', data.taxpayer.address],
      ['4. Nature of Business', ':', data.taxpayer.businessNature],
      ['5. Economic Activity', ':', data.taxpayer.activity]
    ];
    s1Data.forEach(item => {
      const r = sheet.addRow([item[0], item[1], item[2]]);
      sheet.mergeCells(`C${r.number}:E${r.number}`); // Merged for full-width data row
    });
    applyBorder(s1Start, sheet.rowCount);

    addHeader("SECTION - 2: RETURN SUBMISSION DATA");
    const s2Start = sheet.rowCount + 1;
    const s2Data = [
      ['1. Tax Period', ':', data.returnSubmission.period],
      ['2. Type of Return', ':', 'A) Main/Original Return (Section 64)   [ X ]'],
      ['', '', 'B) Late Return (section 65)   [   ]'],
      ['', '', 'C) Amend Return (section 66)   [   ]'],
      ['3. Any activities in this Tax Period?', ':', '[ X ] Yes     [   ] No'],
      ['4. Date of Submission', ':', data.returnSubmission.date]
    ];
    s2Data.forEach(item => {
      const r = sheet.addRow([item[0], item[1], item[2]]);
      sheet.mergeCells(`C${r.number}:E${r.number}`); // Merged for full-width data row
    });
    applyBorder(s2Start, sheet.rowCount);

    // --- SECTION 3 & 4 ---
    addHeader("SECTION - 3: SUPPLY - OUTPUT TAX");
    const s3Start = sheet.rowCount + 1;
    // Header Row
    const head = sheet.addRow(['Nature of Supply', '', 'Note', 'Value (a)', 'SD (b)', 'VAT (c)', '']);
    sheet.mergeCells(`A${head.number}:B${head.number}`);
    head.font = { bold: true };

    // Data Rows (Standard Rated & Total)
    sheet.addRow(['Zero Rated Goods/Service', 'Direct Export', '1', '', '', '', 'Sub form']);
    sheet.addRow(['', 'Deemed Export', '2', '', '', '', 'Sub form']);
    sheet.addRow(['Exempted Goods/Service', '', '3', '', '', '', 'Sub form']);

    // Note 4
    const n4 = sheet.addRow(['Standard Rated Goods/Service', '', '4', 159270.30, 0, 23890.55, 'Sub form']);
    sheet.mergeCells(`A${n4.number}:B${n4.number}`);

    sheet.addRow(['Goods Based on MRP', '', '5', '', '', '', 'Sub form']);
    sheet.addRow(['Goods/Service Based on Specific VAT', '', '6', '', '', '', 'Sub form']);
    sheet.addRow(['Goods/Service Other than Standard Rate', '', '7', '', '', '', 'Sub form']);
    sheet.addRow(['Retail/Whole Sale/Trade Based Supply', '', '8', 0, 0, 0, 'Sub form']);

    // Note 9 Total
    const n9 = sheet.addRow(['Total Sales Value & Total Payable Taxes', '', '9', 159270.30, 0, 23890.55, '']);
    sheet.mergeCells(`A${n9.number}:B${n9.number}`);
    n9.eachCell(c => {
      c.font = { bold: true };
      if (c.address.includes('D') || c.address.includes('E') || c.address.includes('F')) {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9D9D9' } };
      }
    });
    applyBorder(s3Start, sheet.rowCount);

    addHeader("SECTION - 4: PURCHASE - INPUT TAX");
    const s4Start = sheet.rowCount + 1;
    sheet.addRow(['Nature', 'Note', 'Value', 'VAT', 'Remarks']).font = { bold: true };
    sheet.addRow(['Local Purchase', '14', data.notes.note14.val, data.notes.note14.vat, 'Sub form']);
    applyBorder(s4Start, sheet.rowCount);

    // --- INDIVIDUAL SECTIONS 5, 6, 7, 8 (FIXED) ---
    addHeader("SECTION - 5: INCREASING ADJUSTMENTS");
    const s5Start = sheet.rowCount + 1;
    sheet.addRow(['Total Increasing Adjustment', '28', '', data.notes.note28]);
    applyBorder(s5Start, sheet.rowCount);

    addHeader("SECTION - 6: DECREASING ADJUSTMENTS");
    const s6Start = sheet.rowCount + 1;
    sheet.addRow(['Total Decreasing Adjustment', '33', '', data.notes.note33]);
    applyBorder(s6Start, sheet.rowCount);

    addHeader("SECTION - 7: NET TAX CALCULATION");
    const s7Start = sheet.rowCount + 1;
    sheet.addRow(['Net Payable VAT (34)', '34', '', data.notes.note34]);
    sheet.addRow(['Payable for Treasury Deposit', '50', '', data.notes.note50]);
    applyBorder(s7Start, sheet.rowCount);

    addHeader("SECTION - 8: OLD ACCOUNT BALANCE");
    const s8Start = sheet.rowCount + 1;
    sheet.addRow(['Balance from Mushak-18.6', '54', '', data.notes.note54]);
    applyBorder(s8Start, sheet.rowCount);

    // --- INDIVIDUAL SECTIONS 9, 10, 11, 12 ---
    addHeader("SECTION - 9: ACCOUNT CODE WISE PAYMENT SCHEDULE");
    const s9Start = sheet.rowCount + 1;
    sheet.addRow(['VAT Deposit', '58', '1/1133/0030/0311', '0.00']);
    applyBorder(s9Start, sheet.rowCount);

    addHeader("SECTION - 10: CLOSING BALANCE");
    const s10Start = sheet.rowCount + 1;
    sheet.addRow(['Closing Balance (VAT)', '65', '', data.notes.note65]);
    applyBorder(s10Start, sheet.rowCount);

    addHeader("SECTION - 11: REFUND");
    const s11Start = sheet.rowCount + 1;
    sheet.addRow(['Requested Refund (VAT)', '67', '', data.notes.note67]);
    applyBorder(s11Start, sheet.rowCount);

    addHeader("SECTION - 12: DECLARATION");
    const s12Start = sheet.rowCount + 1;
    const dec = sheet.addRow(['I hereby declare that all information are true & accurate.']);
    sheet.mergeCells(`A${dec.number}:E${dec.number}`);
    sheet.addRow(['Name', '', 'Hasanuzzaman']);
    sheet.addRow(['Signature', '', '']);
    applyBorder(s12Start, sheet.rowCount);

    // --- SAVE ---
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), 'Mushak_9.1_Full_Report.xlsx');
  }

}