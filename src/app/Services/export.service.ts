import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

// pdfmake imports
import pdfMakeImport from 'pdfmake/build/pdfmake';
import pdfFontsImport from 'pdfmake/build/vfs_fonts';
import { MushakData } from './MushakData';
import { forkJoin, map, Observable } from 'rxjs';
import { HttpClient } from '@angular/common/http';

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
  constructor(private http: HttpClient) { }

  getMergedMushakData(apiEndpoint: string, lang: 'en' | 'bl'): Observable<any> {
    const labels$ = this.http.get<any>(`i18n/${lang}/dummyData.json`);
    const values$ = this.http.get<any>(apiEndpoint);

    return forkJoin([labels$, values$]).pipe(
      map(([labels, values]) => {
        return {
          labels: labels, // সম্পূর্ণ JSON লেবেলগুলো
          values: values.taxpayer, // এপিআই থেকে আসা মানগুলো
          notes: values.returnSubmission,
          taxpayer: values.taxpayer || labels.taxpayer,
          returnSubmission: this.mergeNotes(labels.notes, values.notes)
        };
      })
    );
  }

  private mergeNotes(labelNotes: any, valueNotes: any): any {
    const merged: any = {};
    Object.keys(labelNotes).forEach(key => {
      merged[key] = {
        label: labelNotes[key],
        val: valueNotes?.[key]?.val ?? 0,
        vat: valueNotes?.[key]?.vat ?? 0,
        sd: valueNotes?.[key]?.sd ?? 0
      };
    });
    return merged;
  }

  // getMushakJsonData(): Observable<MushakData> {
  //   debugger
  //   return this.http.get<MushakData>('i18n/en/dummyData.json');
  // }


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
    debugger
    const l = data.labels || {};
    const n = data?.notes || {};
    const t = data?.taxpayer || {};
    const s = data?.returnSubmission || {};

    (pdfMake as any).fonts = {
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-Regular.ttf'
      }
    };

    const formatAmount = (val: any) => {
      const num = parseFloat(val) || 0;
      return num < 0 ? `(${Math.abs(num)})` : num.toFixed(2);
    };

    const docDef: any = {
      pageSize: 'Legal',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: {
        font: 'PlaywriteCU',
        fontSize: 7
      },
      content: [
        {
          stack: [{ text: l.titles.gov, style: 'header' },
          { text: l.titles.nbr, style: 'header' },
          { text: `\n${l.titles.form}`, style: 'subHeader' },
          { text: `${l.titles.rule}\n`, style: 'subHeader' },
          { text: "\n", style: 'subHeader' }]
        },

        this.createFullWidthHeader(l.sections.s1),
        {
          style: 'dataTable',
          table: {
            widths: ['40%', '2%', '58%'],
            body: [
              [l.labels.bin, ':', t.bin],
              [l.labels.name, ':', t.name],
              [l.labels.address, ':', t.address || ''],
              [l.labels.nature, ':', t.businessNature],
              [l.labels.activity, ':', t.activity]
            ]
          }
        },

        this.createFullWidthHeader(l.sections.s2),
        {
          style: 'dataTable',
          table: {
            widths: ['40%', '2%', '58%'],
            body: [
              [l.labels.tax_period, ':', { text: s.period || 'Oct / 2022', alignment: 'center' }],
              [l.labels.return_type, ':', {
                stack: [
                  { columns: [{ width: '70%', text: l.return_options ? l.return_options[0] : '' }, { table: { widths: ['35%'], body: [[' ']] }, margin: [0, 0, 10, 5], alignment: 'right' }] },
                  { columns: [{ width: '70%', text: l.return_options ? l.return_options[1] : '' }, { table: { widths: ['35%'], body: [[' ']] }, margin: [0, 0, 10, 5], alignment: 'right' }] },
                  { columns: [{ width: '70%', text: l.return_options ? l.return_options[2] : '' }, { table: { widths: ['35%'], body: [[' ']] }, margin: [0, 0, 10, 5], alignment: 'right' }] },
                  { columns: [{ width: '70%', text: l.return_options ? l.return_options[3] : '' }, { table: { widths: ['35%'], body: [[' ']] }, margin: [0, 0, 10, 5], alignment: 'right' }] }
                ], margin: [0, 5, 0, 5]
              }],
              // Row 3: Any activities in this Tax Period?
              [
                l.labels.any_activities,
                ':',
                {
                  stack: [
                    {
                      alignment: 'center',
                      columns: [
                        { width: '*', text: '' },
                        {
                          width: 'auto',
                          columns: [
                            // Yes Option
                            { width: 'auto', table: { widths: [20], body: [[' ']] }, margin: [0, 0, 5, 0] },
                            { width: 'auto', text: l.labels.yes, fontSize: 7, margin: [0, 2, 25, 0] },

                            // No Option
                            { width: 'auto', table: { widths: [20], body: [[' ']] }, margin: [0, 0, 5, 0] },
                            { width: 'auto', text: l.labels.no, fontSize: 7, margin: [0, 2, 0, 0] }
                          ]
                        },
                        { width: '*', text: '' }
                      ]
                    },
                    {
                      text: l.labels.activity_note,
                      fontSize: 7,
                      alignment: 'center',
                      margin: [0, 5, 0, 0],
                      color: '#333333'
                    }
                  ],
                  margin: [0, 10, 0, 10]
                }
              ],
              [l.labels.sub_date, ':', { text: s.date || 'Oct / 2022', alignment: 'center' }]
            ]
          }
        },

        this.createFullWidthHeader(l.sections.s3),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['35%', '10%', '5%', '15%', '13%', '13%', '9%'],
            body: [
              // Table Header
              [
                { text: l.headers.nature_supply, style: 'tHead', colSpan: 2, alignment: 'center' },
                {},
                { text: l.headers.note, style: 'tHead', alignment: 'center' },
                { text: l.headers.value, style: 'tHead', alignment: 'center' },
                { text: l.headers.sd, style: 'tHead', alignment: 'center' },
                { text: l.headers.vat, style: 'tHead', alignment: 'center' },
                { text: '', border: [false, false, false, false] }
              ],
              // Note 1 & 2: Zero Rated
              [
                { text: l.notes.note1.split('-')[0], rowSpan: 2 },
                l.notes.note1.split('-')[1] || l.notes.note1, '1',
                (n['note1']?.val || 0).toLocaleString(),
                (n['note1']?.sd || 0).toLocaleString(),
                (n['note1']?.vat || 0).toLocaleString(),
                l.headers.sub_form
              ],
              [
                {}, l.notes.note2.split('-')[1] || l.notes.note2, '2',
                (n['note2']?.val || 0).toLocaleString(),
                (n['note2']?.sd || 0).toLocaleString(),
                (n['note2']?.vat || 0).toLocaleString(),
                l.headers.sub_form
              ],
              // Note 3: Exempted
              [{ text: l.notes.note3, colSpan: 2 }, {}, '3', (n['note3']?.val || 0).toLocaleString(), '0.00', '0.00', l.headers.sub_form],
              // Note 4: Standard Rated
              [
                { text: l.notes.note4, colSpan: 2 }, {}, '4',
                (n['note4']?.val || 0).toLocaleString(undefined, { minimumFractionDigits: 2 }),
                (n['note4']?.sd || 0).toLocaleString(),
                (n['note4']?.vat || 0).toLocaleString(undefined, { minimumFractionDigits: 2 }),
                l.headers.sub_form
              ],
              // Note 5-8: Other Categories
              [{ text: l.notes.note5, colSpan: 2 }, {}, '5', (n['note5']?.val || 0).toLocaleString(), '0.00', '0.00', l.headers.sub_form],
              [{ text: l.notes.note6, colSpan: 2 }, {}, '6', (n['note6']?.val || 0).toLocaleString(), '0.00', '0.00', l.headers.sub_form],
              [{ text: l.notes.note7, colSpan: 2 }, {}, '7', (n['note7']?.val || 0).toLocaleString(), '0.00', '0.00', l.headers.sub_form],
              [{ text: l.notes.note8, colSpan: 2 }, {}, '8', (n['note8']?.val || 0).toLocaleString(), '0.00', '0.00', l.headers.sub_form],
              // Note 9: Total
              [
                { text: l.notes.note9, colSpan: 2, style: 'tBold' }, {},
                { text: '9', style: 'tBold' },
                { text: (n['note9']?.val || 0).toLocaleString(undefined, { minimumFractionDigits: 2 }), style: 'tBold', fillColor: '#d9d9d9' },
                { text: (n['note9']?.sd || 0).toLocaleString(), style: 'tBold', fillColor: '#d9d9d9' },
                { text: (n['note9']?.vat || 0).toLocaleString(undefined, { minimumFractionDigits: 2 }), style: 'tBold', fillColor: '#d9d9d9' },
                { text: '', border: [false, false, false, false] }
              ]
            ]
          }
        },

        this.createFullWidthHeader(l.sections.s4),
        {
          stack: [
            {
              canvas: [{ type: 'rect', x: 0, y: 0, w: 535, h: 45, color: '#fcd5b4' }]
            },
            {
              text: l.labels.purchase_instruction.join('\n'),
              fontSize: 7,
              margin: [5, -43, 5, 2]
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
                { text: l.labels.nature_purchase, style: 'tHead', colSpan: 2, alignment: 'center' },
                {},
                { text: l.headers.note, style: 'tHead', alignment: 'center' },
                { text: l.headers.value, style: 'tHead', alignment: 'center' },
                { text: l.headers.vat, style: 'tHead', alignment: 'center' },
                { text: '', border: [false, false, false, false] }
              ],
              // Zero Rated & Exempted (Notes 10-13)
              [{ text: l.notes.note10, rowSpan: 2 }, l.labels.local_purchase, '10', n.note10?.val || '0.00', { text: '', fillColor: '#d9d9d9' }, l.headers.sub_form],
              [{}, 'Import', '11', n.note11?.val || '0.00', { text: '', fillColor: '#d9d9d9' }, l.headers.sub_form],
              [{ text: l.notes.note11, rowSpan: 2 }, l.labels.local_purchase, '12', n.note12?.val || '0.00', { text: '', fillColor: '#d9d9d9' }, l.headers.sub_form],
              [{}, 'Import', '13', n.note13?.val || '0.00', { text: '', fillColor: '#d9d9d9' }, l.headers.sub_form],

              // Standard Rated - Main Data (Notes 14-15)
              [{ text: l.notes.note12, rowSpan: 2 }, l.labels.local_purchase, '14', n.note14?.val || '0.00', n.note14?.vat || '0.00', l.headers.sub_form],
              [{}, 'Import', '15', n.note15?.val || '0.00', n.note15?.vat || '0.00', l.headers.sub_form],

              // Other Categories (Notes 16-22)
              [{ text: l.notes.note13, rowSpan: 2 }, l.labels.local_purchase, '16', n.note16?.val || '0.00', n.note16?.vat || '0.00', l.headers.sub_form],
              [{}, 'Import', '17', n.note17?.val || '0.00', n.note17?.vat || '0.00', l.headers.sub_form],
              [{ text: l.notes.note14, rowSpan: 1 }, l.labels.local_purchase, '18', n.note18?.val || '0.00', n.note18?.vat || '0.00', l.headers.sub_form],
              [{ text: l.notes.note15, rowSpan: 2 }, 'From Turnover Units', '19', n.note19?.val || '0.00', n.note19?.vat || '0.00', l.headers.sub_form],
              [{}, 'From Unregistered Entities', '20', n.note20?.val || '0.00', n.note20?.vat || '0.00', l.headers.sub_form],
              [{ text: l.notes.note16, rowSpan: 2 }, l.labels.local_purchase, '21', n.note21?.val || '0.00', n.note21?.vat || '0.00', l.headers.sub_form],
              [{}, 'Import', '22', n.note22?.val || n.note22?.val || '0.00', n.note22?.vat || '0.00', l.headers.sub_form],

              // Total Row (Note 23)
              [
                { text: l.labels.total_input_credit, colSpan: 1, style: 'tBold', bold: true },
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
        this.createFullWidthHeader(l.sections.s5),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['50%', '10%', '25%', '15%'],
            body: [
              [
                { text: l.headers.adj_details, style: 'tHead', alignment: 'center' },
                { text: l.headers.note, style: 'tHead', alignment: 'center' },
                { text: l.headers.vat_amount, style: 'tHead', alignment: 'center' },
                { text: '', style: 'tHead', border: [false, false, false, false] }
              ],
              // Note 24-26
              [l.notes.note24, { text: '24', alignment: 'center' }, { text: '0.00', alignment: 'right' }, l.headers.sub_form],
              [l.notes.note25, { text: '25', alignment: 'center' }, { text: '0.00', alignment: 'right' }, l.headers.sub_form],
              [l.notes.note26, { text: '26', alignment: 'center' }, '', l.headers.sub_form],
              // Note 27: Other Adjustments with Stacked Label
              [
                {
                  stack: [
                    l.notes.note27,
                    {
                      margin: [0, 5, 0, 0],
                      table: {
                        width: '*',
                        body: [[{ text: l.notes.note27_sub, fontSize: 7, bold: true }]]
                      },
                    },
                    // { text: 'VAT on House Rent', margin: [0, 5, 0, 0], bold: true }
                  ]
                },
                { text: '27', alignment: 'center' },
                { text: n.note27?.val || '0.00', alignment: 'right' },
                l.headers.sub_form
              ],
              // Row 5: Total (Note 28)
              [
                { text: l.labels.total_inc_adj, style: 'tBold', bold: true },
                { text: '28', style: 'tBold', alignment: 'center' },
                { text: n.note28?.val || n.note28 || '0.00', style: 'tBold', alignment: 'right' },
                { text: '', border: [false, false, false, false] }
              ]
            ]
          }
        },

        // Inside exportFullMushakPdf
        this.createFullWidthHeader(l.sections.s6),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['50%', '10%', '25%', '15%'],
            body: [
              [
                { text: l.headers.adj_details, style: 'tHead', alignment: 'center' },
                { text: l.headers.note, style: 'tHead', alignment: 'center' },
                { text: l.headers.vat_amount, style: 'tHead', alignment: 'center' },
                { text: '', style: 'tHead', border: [false, false, false, false] }
              ],
              // Note 29: VDS from supplies delivered
              [l.notes.note29, { text: '29', alignment: 'center' }, { text: n.note29?.val || '0.00', alignment: 'right' }, l.headers.sub_form],
              // Note 30: Advance Tax
              [l.notes.note30, { text: '30', alignment: 'center' }, { text: n.note30?.val || '0.00', alignment: 'right' }, l.headers.sub_form],
              // Note 31: Credit Note
              [l.notes.note31, { text: '31', alignment: 'center' }, { text: n.note31?.val || '0.00', alignment: 'right' }, l.headers.sub_form],
              // Note 32: Other Adjustments with empty box
              [
                {
                  stack: [
                    l.notes.note32,
                    {
                      table: { widths: ['*'], body: [[' ']] },
                      margin: [0, 5, 10, 2]
                    }
                  ]
                },
                { text: '32', alignment: 'center' },
                { text: n.note32?.val || '0.00', alignment: 'right' },
                l.headers.sub_form
              ],
              // Row 5: Total Decreasing Adjustment (Note 33)
              [
                { text: l.labels.total_dec_adj, style: 'tBold', bold: true },
                { text: '33', style: 'tBold', alignment: 'center' },
                { text: (n.note33 || '0.00'), style: 'tBold', alignment: 'right' },
                { text: '', border: [false, false, false, false] }
              ]
            ]
          }
        },

        this.createFullWidthHeader(l.sections.s7),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['*', 40, 100], // Exactly 3 columns defined
            body: [
              // Row 0: Header
              [
                { text: l.headers.items, style: 'tHead', alignment: 'center' },
                { text: l.headers.note, style: 'tHead', alignment: 'center' },
                { text: l.headers.amount, style: 'tHead', alignment: 'center' }
              ],
              // Notes 34 - 53
              [l.notes.note34, '34', { text: formatAmount(n.note34?.val || n.note34), alignment: 'right' }],
              [l.notes.note35, '35', { text: formatAmount(n.note35?.val || n.note35), alignment: 'right' }],
              [l.notes.note36, '36', { text: formatAmount(n.note36?.val || n.note36), alignment: 'right' }],
              [l.notes.note37, '37', { text: formatAmount(n.note37?.val || n.note37), alignment: 'right' }],
              [l.notes.note38, '38', { text: formatAmount(n.note38?.val || n.note38), alignment: 'right' }],
              [l.notes.note39, '39', { text: formatAmount(n.note39?.val || n.note39), alignment: 'right' }],
              [l.notes.note40, '40', { text: formatAmount(n.note40?.val || n.note40), alignment: 'right' }],
              [l.notes.note41, '41', { text: formatAmount(n.note41?.val || n.note41), alignment: 'right' }],
              [l.notes.note42, '42', { text: formatAmount(n.note42?.val || n.note42), alignment: 'right' }],
              [l.notes.note43, '43', { text: formatAmount(n.note43?.val || n.note43), alignment: 'right' }],
              [l.notes.note44, '44', { text: formatAmount(n.note44?.val || n.note44), alignment: 'right' }],
              [l.notes.note45, '45', { text: formatAmount(n.note45?.val || n.note45), alignment: 'right' }],
              [l.notes.note46, '46', { text: formatAmount(n.note46?.val || n.note46), alignment: 'right' }],
              [l.notes.note47, '47', { text: formatAmount(n.note47?.val || n.note47), alignment: 'right' }],
              [l.notes.note48, '48', { text: formatAmount(n.note48?.val || n.note48), alignment: 'right' }],
              [l.notes.note49, '49', { text: formatAmount(n.note49?.val || n.note49), alignment: 'right' }],
              [l.notes.note50, '50', { text: formatAmount(n.note50?.val || n.note50), alignment: 'right', style: 'tBold' }],
              [l.notes.note51, '51', { text: formatAmount(n.note51?.val || n.note51), alignment: 'right' }],
              [l.notes.note52, '52', { text: formatAmount(n.note52?.val || n.note52), alignment: 'right', style: 'tBold' }],
              [l.notes.note53, '53', { text: formatAmount(n.note53?.val || n.note53), alignment: 'right' }]
            ]
          }
        },

        this.createFullWidthHeader(l.sections.s8),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['*', 40, 100], // Matches Section 7's stable 3-column layout
            body: [
              // Row 0: Header
              [
                { text: l.headers.items, style: 'tHead', alignment: 'center' },
                { text: l.headers.note, style: 'tHead', alignment: 'center' },
                { text: l.headers.amount, style: 'tHead', alignment: 'center' }
              ],
              // Notes 54 - 57
              [
                l.notes.note54,
                { text: '54', alignment: 'center' },
                { text: (n.note54 || '0.00').toLocaleString(), alignment: 'right' }
              ],
              [
                l.notes.note55,
                { text: '55', alignment: 'center' },
                { text: (n.note55?.val || n.note55 || '0.00'), alignment: 'right' }
              ],
              [
                l.notes.note56,
                { text: '56', alignment: 'center' },
                { text: (n.note56?.val || n.note56 || '0.00'), alignment: 'right' }
              ],
              [
                l.notes.note57,
                { text: '57', alignment: 'center' },
                { text: (n.note57?.val || n.note57 || '0.00'), alignment: 'right' }
              ]
            ]
          }
        },

        { text: '', pageBreak: 'before' },
        this.createFullWidthHeader(l.sections.s9),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['35%', '10%', '25%', '18%', '12%'],
            body: [
              [
                { text: l.headers.items, style: 'tHead', alignment: 'center' },
                { text: l.headers.note, style: 'tHead', alignment: 'center' },
                { text: l.headers.acc_code, style: 'tHead', alignment: 'center' },
                { text: l.headers.amount, style: 'tHead', alignment: 'center' },
                { text: '', style: 'tHead' }
              ],
              // Row 58: VAT Deposit
              [l.notes.note58, '58', n.note58?.code || '1/1133/0030/0311', n.note58?.val || '0.00', l.headers.sub_form],
              // Row 59: SD Deposit
              [l.notes.note59, '59', n.note59?.code || '1/1133/0018/0711-0721', n.note59?.val || '0.00', l.headers.sub_form],
              // Row 60: Excise Duty
              [l.notes.note60, '60', n.note60?.code || '1/1133/Acv‡ikbvj †KvW/0311', n.note60?.val || '0.00', l.headers.sub_form],
              // Row 61: Development Surcharge
              [l.notes.note61, '61', n.note61?.code || '1/1133/Acv‡ikbvj', n.note61?.val || '0.00', l.headers.sub_form],
              // Row 62: ICT Development Surcharge
              [l.notes.note62, '62', n.note62?.code || '1/1103/Acv‡ikbvj †KvW/1901', n.note62?.val || '0.00', l.headers.sub_form],
              // Row 63: Health Care Surcharge
              [l.notes.note63, '63', n.note63?.code || '1/1133/Acv‡ikbvj †KvW/0601', n.note63?.val || '0.00', l.headers.sub_form],
              // Row 64: Environmental Protection Surcharge
              [l.notes.note64, '64', n.note64?.code || '1/1103/Acv‡ikbvj †KvW/2225', n.note64?.val || '0.00', l.headers.sub_form]
            ]
          }
        },

        this.createFullWidthHeader(l.sections.s10),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['*', 40, 100], // Stable 3-column layout
            body: [
              // Row 0: Header
              [
                { text: l.headers.items, style: 'tHead', alignment: 'center' },
                { text: l.headers.note, style: 'tHead', alignment: 'center' },
                { text: l.headers.amount, style: 'tHead', alignment: 'center' }
              ],
              // Row 65: Closing Balance (VAT)
              [
                l.notes.note65,
                { text: '65', alignment: 'center' },
                { text: (n.note65?.val || n.note65 || '0.00'), alignment: 'right', style: 'tBold' }
              ],
              // Row 66: Closing Balance (SD)
              [
                l.notes.note66,
                { text: '66', alignment: 'center' },
                { text: (n.note66?.val || n.note66 || '0.00'), alignment: 'right' }
              ]
            ]
          }
        },

        this.createFullWidthHeader(l.sections.s11),
        {
          style: 'dataTable',
          table: {
            widths: ['35%', '35%', '10%', '20%'], // 4-column layout
            body: [
              // Header Row
              [
                { text: l.labels.interest_refund, rowSpan: 3, margin: [0, 10] },
                { text: l.headers.items, style: 'tHead', alignment: 'center' },
                { text: l.headers.note, style: 'tHead', alignment: 'center' },
                {
                  columns: [
                    // { text: l.labels.yes, fontSize: 7 }, { text: '[  ]', fontSize: 7 },
                    // { text: l.labels.no, fontSize: 7 }, { text: '[  ]', fontSize: 7 }
                    { width: 'auto', table: { widths: [20], body: [[' ']] }, margin: [0, 0, 5, 0] },
                    { width: 'auto', text: l.labels.yes, fontSize: 7, margin: [0, 2, 25, 0] },

                    // No Option
                    { width: 'auto', table: { widths: [20], body: [[' ']] }, margin: [0, 0, 5, 0] },
                    { width: 'auto', text: l.labels.no, fontSize: 7, margin: [0, 2, 0, 0] }
                  ],
                  style: 'tHead'
                }
              ],
              // Note 67
              [
                {},
                l.labels.req_refund_vat,
                { text: '67', alignment: 'center' },
                { text: (n.note67?.val || n.note67 || '0.00'), alignment: 'right' }
              ],
              // Note 68
              [
                {},
                l.labels.req_refund_sd,
                { text: '68', alignment: 'center' },
                { text: (n.note68?.val || n.note68 || '0.00'), alignment: 'right' }
              ]
            ]
          }
        },

        this.createFullWidthHeader(l.sections.s12),
        {
          style: 'dataTable',
          margin: [0, 0, 0, 0],
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: l.labels.declaration_text,
                  fillColor: '#d9d9d9', // Gray background for the disclaimer
                  fontSize: 7,
                  margin: [5, 5, 5, 5]
                }
              ]
            ],
          }
        },
        {
          style: 'dataTable',
          table: {
            widths: ['35%', '5%', '60%'], // Matches the alignment of Section 1
            body: [
              [l.labels.name, ':', t.name || ''],
              [l.labels.designation, ':', t.designation || ''],
              [l.labels.mobile, ':', t.mobile || ''],
              [l.labels.nid_passport, ':', t.nid_passport || ''],
              [l.labels.email, ':', t.email || ''],
              [l.labels.signature, ':', '']
            ]
          }
        }
      ],
      styles: {
        header: { font: 'PlaywriteCU', fontSize: 8, bold: true, alignment: 'center' },
        subHeader: { font: 'PlaywriteCU', fontSize: 7, bold: true, alignment: 'center', color: '#003366' },
        secHeaderCell: { font: 'PlaywriteCU', fillColor: '#003366', color: 'white', bold: true, alignment: 'center', fontSize: 7, padding: [0, 2, 0, 2] },
        tHead: { font: 'PlaywriteCU', fillColor: '#f2f2f2', bold: true, fontSize: 7 },
        tBold: { font: 'PlaywriteCU', bold: true, fontSize: 7 },
        dataTable: { font: 'PlaywriteCU', fontSize: 7, margin: [0, 0, 0, 5] },
        borderedTable: { font: 'PlaywriteCU', margin: [0, 0, 0, 2] }
      }
    };
    pdfMake.createPdf(docDef).download('Mushak_9.1_Full_Report.pdf');
  }

  // ফন্টের নাম অনুযায়ী স্টাইল জেনারেট করার ফাংশন
private getMushakStyles(fontName: string) {
  return {
    header: { 
      font: fontName, 
      fontSize: 8, 
      bold: true, 
      alignment: 'center' 
    },
    subHeader: { 
      font: fontName, 
      fontSize: 7, 
      bold: true, 
      alignment: 'center', 
      color: '#003366' 
    },
    secHeaderCell: { 
      font: fontName, 
      fillColor: '#003366', 
      color: 'white', 
      bold: true, 
      alignment: 'center', 
      fontSize: 7, 
      padding: [0, 2, 0, 2] 
    },
    tHead: { 
      font: fontName, 
      fillColor: '#f2f2f2', 
      bold: true, 
      fontSize: 7 
    },
    tBold: { 
      font: fontName, 
      bold: true, 
      fontSize: 7 
    },
    dataTable: { 
      font: fontName, 
      fontSize: 7, 
      margin: [0, 0, 0, 5] 
    },
    borderedTable: { 
      font: fontName, 
      margin: [0, 0, 0, 2] 
    }
  };
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
      { width: 12 }  // F: সাবফর্ম
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
    sheet.addRow(['Zero Rated Goods/Service', 'Direct Export', '1', '', '', '', 'সাবফর্ম']);
    sheet.addRow(['', 'Deemed Export', '2', '', '', '', 'সাবফর্ম']);
    sheet.addRow(['Exempted Goods/Service', '', '3', '', '', '', 'সাবফর্ম']);

    // Note 4
    const n4 = sheet.addRow(['Standard Rated Goods/Service', '', '4', 159270.30, 0, 23890.55, 'সাবফর্ম']);
    sheet.mergeCells(`A${n4.number}:B${n4.number}`);

    sheet.addRow(['Goods Based on MRP', '', '5', '', '', '', 'সাবফর্ম']);
    sheet.addRow(['Goods/Service Based on Specific VAT', '', '6', '', '', '', 'সাবফর্ম']);
    sheet.addRow(['Goods/Service Other than Standard Rate', '', '7', '', '', '', 'সাবফর্ম']);
    sheet.addRow(['Retail/Whole Sale/Trade Based Supply', '', '8', 0, 0, 0, 'সাবফর্ম']);

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
    sheet.addRow(['Local Purchase', '14', data.notes.note14.val, data.notes.note14.vat, 'সাবফর্ম']);
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


  exportFullMushakPdfBangla(data: any) {
    const n = data?.notes || {};
    const t = data?.taxpayer || {};
    const s = data?.returnSubmission || {};

    (pdfMake as any).fonts = {
      BanglaFonts: {
        normal: window.location.origin + '/assets/fonts/sagarnormal.ttf',
        bold: window.location.origin + '/assets/fonts/sagarnormal.ttf',
        italics: window.location.origin + '/assets/fonts/sagarnormal.ttf',
        bolditalics: window.location.origin + '/assets/fonts/sagarnormal.ttf'
      }
    };
    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: {
        font: 'BanglaFonts',
        fontSize: 7
      },
      content: [
        {
          stack: [{ text: "গণপ্রজাতন্ত্রী বাংলাদেশ সরকার", style: 'header' },
          { text: "জাতীয় রাজস্ব বোর্ড", style: 'header' },
          { text: "\nমূল্য সংযোজন কর দাখিলপত্র (মূসক-৯.১)", style: 'subHeader' },
          { text: "\n[বিধি ৪৭(১) দ্রষ্টব্য]", style: 'subHeader' },
          { text: "\n", style: 'subHeader' }]
        },

        this.createFullWidthHeader("অংশ - ১: করদাতার তথ্য"),
        {
          style: 'dataTable',
          table: {
            widths: ['40%', '2%', '58%'],
            body: [
              ['১. বিআইএন (BIN)', ':', t.bin],
              ['২. নাম', ':', t.name],
              ['৩. ঠিকানা', ':', t.address || ''],
              ['৪. প্রকার', ':', t.businessNature],
              ['৫. কার্যক্রম', ':', t.activity]
            ]
          }
        },

        this.createFullWidthHeader("অংশ - ২: দাখিলপত্র পেশের তথ্য"),
        {
          style: 'dataTable',
          table: {
            widths: ['40%', '2%', '58%'],
            body: [
              ['১. কর মেয়াদ', ':', { text: s.period || 'অক্টোবর / ২০২২', alignment: 'center' }],
              ['২. দাখিলপত্রের ধরন\n[অনুগ্রহ করে আপনার কাঙ্ক্ষিত অপশনটি নির্বাচন করুন]', ':', {
                stack: [
                  { columns: [{ width: '70%', text: 'ক) মূল দাখিলপত্র (ধারা ৬৪)' }, { table: { widths: ['35%'], body: [[' ']] }, margin: [0, 0, 10, 5], alignment: 'right' }] },
                  { columns: [{ width: '70%', text: 'খ) বিলম্বিত দাখিলপত্র (ধারা ৬৫)' }, { table: { widths: ['35%'], body: [[' ']] }, margin: [0, 0, 10, 5], alignment: 'right' }] },
                  { columns: [{ width: '70%', text: 'গ) সংশোধিত দাখিলপত্র (ধারা ৬৬)' }, { table: { widths: ['35%'], body: [[' ']] }, margin: [0, 0, 10, 5], alignment: 'right' }] },
                  { columns: [{ width: '70%', text: 'ঘ) পূর্ণাঙ্গ/অতিরিক্ত/বিকল্প দাখিলপত্র (ধারা ৬৭)' }, { table: { widths: ['35%'], body: [[' ']] }, margin: [0, 0, 10, 5], alignment: 'right' }] }
                ], margin: [0, 5, 0, 5]
              }],
              // Row 3: Any activities in this Tax Period?
              [
                '৩. এই কর মেয়াদে কি কোনো কার্যক্রম ছিল?',
                ':',
                {
                  stack: [
                    {
                      alignment: 'center',
                      columns: [
                        { width: '*', text: '' },
                        {
                          width: 'auto',
                          columns: [
                            // Yes Option
                            { width: 'auto', table: { widths: [20], body: [[' ']] }, margin: [0, 0, 5, 0] },
                            { width: 'auto', text: 'হ্যাঁ', fontSize: 7, margin: [0, 2, 25, 0] },

                            // No Option
                            { width: 'auto', table: { widths: [20], body: [[' ']] }, margin: [0, 0, 5, 0] },
                            { width: 'auto', text: 'না', fontSize: 7, margin: [0, 2, 0, 0] }
                          ]
                        },
                        { width: '*', text: '' }
                      ]
                    },
                    {
                      text: '[যদি "না" নির্বাচিত হয় তবে অনুগ্রহ করে শুধুমাত্র বিভাগ I, II পূরণ করুন]',
                      fontSize: 7,
                      alignment: 'center',
                      margin: [0, 5, 0, 0],
                      color: '#333333'
                    }
                  ],
                  margin: [0, 10, 0, 10] // Vertical padding for the whole cell
                }
              ],
              ['৪. দাখিলপত্র পেশের তারিখ', ':', { text: s.date || '03-Oct-2022', alignment: 'center' }]
            ]
          }
        },

        this.createFullWidthHeader("অংশ - ৩: সরবরাহ (উৎপাদ কর)"),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['35%', '10%', '5%', '15%', '13%', '13%', '9%'],
            body: [
              // Table Header
              [
                { text: 'সরবরাহের ধরন', style: 'tHead', colSpan: 2, alignment: 'center' },
                {},
                { text: 'নোট', style: 'tHead', alignment: 'center' },
                { text: 'মূল্য (ক)', style: 'tHead', alignment: 'center' },
                { text: 'সম্পূরক শুল্ক (খ)', style: 'tHead', alignment: 'center' },
                { text: 'মূসক (গ)', style: 'tHead', alignment: 'center' },
                { text: '', style: 'tHead', border: [false, false, false, false] }
              ],
              // Row 1 & 2: Zero Rated
              [
                { text: 'শূন্য হারযুক্ত পণ্য/সেবা', rowSpan: 2 },
                'সরাসরি রপ্তানি', '1', '0.00', { text: '', fillColor: '#d9d9d9' }, { text: '', fillColor: '#d9d9d9' }, 'সাবফর্ম'
              ],
              [
                {}, 'প্রচ্ছন্ন রপ্তানি', '2', '0.00', { text: '', fillColor: '#d9d9d9' }, { text: '', fillColor: '#d9d9d9' }, 'সাবফর্ম'
              ],
              // Row 3: Exempted
              [{ text: 'অব্যাহতি প্রাপ্ত পণ্য/সেবা', colSpan: 2 }, {}, '3', '0.00', { text: '', fillColor: '#d9d9d9' }, { text: '', fillColor: '#d9d9d9' }, 'সাবফর্ম'],
              // Row 4: Standard Rated
              [
                { text: 'আদর্শ হার বিশিষ্ট পণ্য/সেবা', colSpan: 2 }, {}, '4',
                (n.note4?.val || 159270.30).toLocaleString(undefined, { minimumFractionDigits: 2 }),
                '0.00',
                (n.note4?.vat || 23890.55).toLocaleString(undefined, { minimumFractionDigits: 2 }),
                'সাবফর্ম'
              ],
              // Rows 5-7
              [{ text: 'সর্বোচ্চ খুচরা মূল্য ভিত্তিক পণ্য', colSpan: 2 }, {}, '5', '', '', '', 'সাবফর্ম'],
              [{ text: 'নির্দিষ্ট ভ্যাটের ভিত্তিতে পণ্য/সেবা', colSpan: 2 }, {}, '6', '', '', '', 'সাবফর্ম'],
              [{ text: 'আদর্শ হার ব্যতীত পণ্য/সেবা', colSpan: 2 }, {}, '7', '', '', '', 'সাবফর্ম'],
              // Row 8
              [{ text: 'খুচরা/পাইকারি/বাণিজ্য ভিত্তিক সরবরাহ', colSpan: 2 }, {}, '8', '0.00', '0.00', '0.00', { text: 'সাবফর্ম', fillColor: '#d9d9d9' }],
              // Row 9: Total
              [
                { text: 'মোট বিক্রয় মূল্য & মোট প্রদেয় করসমূহ', colSpan: 2, style: 'tBold', bold: true }, {},
                { text: '9', style: 'tBold' },
                { text: (n.note9?.val || 159270.30).toLocaleString(undefined, { minimumFractionDigits: 2 }), style: 'tBold', fillColor: '#d9d9d9' },
                { text: '0.00', style: 'tBold', fillColor: '#d9d9d9' },
                { text: (n.note9?.vat || 23890.55).toLocaleString(undefined, { minimumFractionDigits: 2 }), style: 'tBold', fillColor: '#d9d9d9' },
                { text: '', border: [false, false, false, false] }
              ]
            ]
          }
        },

        this.createFullWidthHeader("অংশ - ৪: ক্রয় (উপকরণ কর)"),
        {
          stack: [
            {
              canvas: [{ type: 'rect', x: 0, y: 0, w: 535, h: 45, color: '#fcd5b4' }]
            },
            {
              text: [
                "১) আপনি যে সকল পণ্য/সেবা সরবরাহ করেন সেগুলো যদি সবই স্ট্যান্ডার্ড রেটেড হয়, তবে নোট ১০-২০ পূরণ করুন।\n",
                "২) আপনি যে সকল পণ্য/সেবা সরবরাহ করেন সেগুলো স্ট্যান্ডার্ড রেটেড নয় অথবা ধারা ৪৬ অনুযায়ী নির্ধারিত সময়সীমার মধ্যে ইনপুট ট্যাক্স ক্রেডিট গ্রহণ করা হয়নি, নোট ২১-২২ পূরণ করুন।\n",
                "৩) আপনি যে পণ্য/সেবা সরবরাহ করেন তা যদি স্ট্যান্ডার্ড রেটেড এবং নন-স্ট্যান্ডার্ড রেটেড উভয়ই নিয়ে গঠিত হয়, তবে স্ট্যান্ডার্ড রেটেড পণ্য/সেবা উৎপাদন/সরবরাহে ব্যবহৃত কাঁচামালের জন্য নোট ১০-২০ পূরণ করুন এবং নন-স্ট্যান্ডার্ড রেটেড পণ্য/সেবা উৎপাদন/সরবরাহে ব্যবহৃত কাঁচামালের জন্য নোট ২১-২২ পূরণ করুন এবং প্রযোজ্য ক্ষেত্রে নোট ১০-২২-এ অনুপাতে মূল্য প্রদর্শন করুন।"
              ],
              fontSize: 7,
              margin: [5, -43, 5, 2]
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
                { text: 'ক্রয়ের প্রকৃতি', style: 'tHead', colSpan: 2, alignment: 'center' },
                {},
                { text: 'নোট', style: 'tHead', alignment: 'center' },
                { text: 'মূল্য (ক)', style: 'tHead', alignment: 'center' },
                { text: 'মূসক (খ)', style: 'tHead', alignment: 'center' },
                { text: '', border: [false, false, false, false] }
              ],
              // Zero Rated & Exempted (Notes 10-13)
              [{ text: 'শূন্য হারের পণ্য/সেবা', rowSpan: 2 }, 'স্থানীয় ক্রয়', '10', '0.00', { text: '', fillColor: '#d9d9d9' }, 'সাবফর্ম'],
              [{}, 'আমদানি', '11', '0.00', { text: '', fillColor: '#d9d9d9' }, 'সাবফর্ম'],
              [{ text: 'অব্যাহতি প্রাপ্ত পণ্য/সেবা', rowSpan: 2 }, 'স্থানীয় ক্রয়', '12', '0.00', { text: '', fillColor: '#d9d9d9' }, 'সাবফর্ম'],
              [{}, 'আমদানি', '13', '0.00', { text: '', fillColor: '#d9d9d9' }, 'সাবফর্ম'],

              // Standard Rated - Main Data (Notes 14-15)
              [{ text: 'আদর্শ হার বিশিষ্ট পণ্য/সেবা', rowSpan: 2 }, 'স্থানীয় ক্রয়', '14', (n.note14?.val || 3717678.34).toLocaleString(), (n.note14?.vat || 557651.75).toLocaleString(), 'সাবফর্ম'],
              [{}, 'আমদানি', '15', '0.00', '0.00', 'সাবফর্ম'],

              // Other Categories (Notes 16-22)
              [{ text: 'আদর্শ হার ব্যতীত পণ্য/সেবা', rowSpan: 2 }, 'স্থানীয় ক্রয়', '16', '0.00', '0.00', 'সাবফর্ম'],
              [{}, 'আমদানি', '17', '0.00', '0.00', 'সাবফর্ম'],
              [{ text: 'নির্দিষ্ট ভ্যাটের ভিত্তিতে পণ্য/সেবা', rowSpan: 1 }, 'স্থানীয় ক্রয়', '18', '0.00', '0.00', 'সাবফর্ম'],
              [{ text: 'ক্রেডিটের জন্য গ্রহণযোগ্য নয় এমন পণ্য/সেবা (স্থানীয় ক্রয়)', rowSpan: 2 }, 'টার্নওভার ইউনিট থেকে', '19', '0.00', '0.00', 'সাবফর্ম'],
              [{}, 'অনিবন্ধিত সত্তাসমূহ থেকে', '20', '0.00', '0.00', 'সাবফর্ম'],
              [{ text: 'পণ্য/সেবা ক্রেডিটের জন্য গ্রহণযোগ্য নয় (যেসব করদাতা শুধুমাত্র অব্যাহতিপ্রাপ্ত/ নির্দিষ্ট ভ্যাট এবং আদর্শ হারের বাইরে পণ্য/সেবা বিক্রি করেন/\nক্রেডিট নেওয়া হয়নি\n', rowSpan: 2 }, 'স্থানীয় ক্রয়', '21', '0.00', '0.00', 'সাবফর্ম'],
              [{}, 'আমদানি', '22', '0.00', '0.00', 'সাবফর্ম'],

              // Total Row (Note 23)
              [
                { text: 'মোট উপকরণ কর রেয়াত', colSpan: 1, style: 'tBold', bold: true },
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
        this.createFullWidthHeader("সেকশন - ৫: বর্ধনমূলক সমন্বয়সমূহ (ভ্যাট)"),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['50%', '10%', '25%', '15%'],
            body: [
              [
                { text: 'সমন্বয় বিবরণ', style: 'tHead', alignment: 'center' },
                { text: 'নোট', style: 'tHead', alignment: 'center' },
                { text: 'ভ্যাট পরিমাণ', style: 'tHead', alignment: 'center' },
                { text: '', style: 'tHead', border: [false, false, false, false] }
              ],
              // Note 24-26
              ['উৎস কর কর্তনের কারণে বৃদ্ধিকারী সমন্বয়', { text: '24', alignment: 'center' }, { text: '0.00', alignment: 'right' }, 'সাবফর্ম'],
              ['ব্যাংকিং চ্যানেলে পেমেন্ট না করার কারণে', { text: '25', alignment: 'center' }, { text: '0.00', alignment: 'right' }, 'সাবফর্ম'],
              ['ডেবিট নোট ইস্যুর কারণে', { text: '26', alignment: 'center' }, '', 'সাবফর্ম'],
              // Note 27: Other Adjustments with Stacked Label
              [
                {
                  stack: [
                    'অন্যান্য কোনো সমন্বয় (অনুগ্রহ করে নিচে উল্লেখ করুন)',
                    {
                      margin: [0, 5, 0, 0],
                      table: {
                        width: '*',
                        body: [[{ text: 'বাড়ি ভাড়ার উপর ভ্যাট', fontSize: 7, bold: true }]]
                      },
                    },
                    // { text: 'VAT on House Rent', margin: [0, 5, 0, 0], bold: true }
                  ]
                },
                { text: '27', alignment: 'center' },
                '',
                'সাবফর্ম'
              ],
              // Row 5: Total (Note 28)
              [
                { text: 'সর্বমোট বৃদ্ধিকারী সমন্বয়', style: 'tBold', bold: true },
                { text: '28', style: 'tBold', alignment: 'center' },
                { text: (n.note28 || '0.00'), style: 'tBold', alignment: 'right' },
                { text: '', border: [false, false, false, false] }
              ]
            ]
          }
        },

        // Inside exportFullMushakPdf
        this.createFullWidthHeader("সেকশন - ৬: হ্রাসকারী সমন্বয় (ভ্যাট)"),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['50%', '10%', '25%', '15%'],
            body: [
              [
                { text: 'সমন্বয়ের বিবরণ', style: 'tHead', alignment: 'center' },
                { text: 'নোট', style: 'tHead', alignment: 'center' },
                { text: 'ভ্যাট পরিমাণ', style: 'tHead', alignment: 'center' },
                { text: '', style: 'tHead', border: [false, false, false, false] }
              ],
              // Note 29: VDS from supplies delivered
              ['সরবরাহকৃত সরবরাহসমূহ থেকে উৎসে ভ্যাট কর্তনের কারণে', { text: '29', alignment: 'center' }, { text: '0.00', alignment: 'right' }, 'সাবফর্ম'],
              // Note 30: Advance Tax
              ['আমদানি পর্যায়ে পরিশোধিত অগ্রিম কর', { text: '30', alignment: 'center' }, '', 'সাবফর্ম'],
              // Note 31: Credit Note
              ['ক্রেডিট নোট ইস্যুর কারণে', { text: '31', alignment: 'center' }, { text: '0.00', alignment: 'right' }, 'সাবফর্ম'],
              // Note 32: Other Adjustments with empty box
              [
                {
                  stack: [
                    'অন্যান্য কোনো সমন্বয় (অনুগ্রহ করে নিচে উল্লেখ করুন)',
                    {
                      table: { widths: ['*'], body: [[' ']] },
                      margin: [0, 5, 10, 2]
                    }
                  ]
                },
                { text: '32', alignment: 'center' },
                { text: '0.00', alignment: 'right' },
                'সাবফর্ম'
              ],
              // Row 5: Total Decreasing Adjustment (Note 33)
              [
                { text: 'সর্বমোট হ্রাসকারী সমন্বয়', style: 'tBold', bold: true },
                { text: '33', style: 'tBold', alignment: 'center' },
                { text: (n.note33 || '0.00'), style: 'tBold', alignment: 'right' },
                { text: '', border: [false, false, false, false] }
              ]
            ]
          }
        },

        this.createFullWidthHeader("অংশ - ৭: নিট কর হিসাব"),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['*', 40, 100], // Exactly 3 columns defined
            body: [
              // Row 0: Header
              [
                { text: 'Items', style: 'tHead', alignment: 'center' },
                { text: 'Note', style: 'tHead', alignment: 'center' },
                { text: 'Amount', style: 'tHead', alignment: 'center' }
              ],
              // Notes 34 - 53
              ['কর পর্বের জন্য নেট প্রদেয় ভ্যাট (ধারা- 45) (9C-23B+28-33)', '34', { text: `(${Math.abs(n.note34 || 533761.20).toLocaleString()})`, alignment: 'right' }],
              ['সমাপনী স্থিতি এবং ফর্ম ১৮.৬ এর স্থিতির সাথে সমন্বয়ের পর কর পর্বের জন্য নেট পরিশোধযোগ্য ভ্যাট [34-(52+56)]', '35', { text: `(${Math.abs(n.note35 || 1979177.91).toLocaleString()})`, alignment: 'right' }],
              ['কর সময়ের জন্য নিট পরিশোধযোগ্য সম্পূরক শুল্ক (সমাপনী ব্যালেন্সের সাথে সমন্বয়ের পূর্বে) [9B+38-(39+40)]', '36', { text: '0.00', alignment: 'right' }],
              ['সমাপনী স্থিতি এবং ফর্ম ১৮.৬-এর অবশিষ্টের সাথে সমন্বয়ের পর কর পর্বের জন্য নিট প্রদেয় সম্পূরক শুল্ক [36-(53+57)', '37', { text: '0.00', alignment: 'right' }],
              ['ডেবিট নোট ইস্যুর বিপরীতে সম্পূরক শুল্ক', '38', { text: '0.00', alignment: 'right' }],
              ['ক্রেডিট নোট ইস্যুর বিপরীতে সম্পূরক শুল্ক', '39', { text: '0.00', alignment: 'right' }],
              ['রপ্তানির বিপরীতে ইনপুটের উপর প্রদত্ত সম্পূরক শুল্ক', '40', { text: '0.00', alignment: 'right' }],
              ['বিলম্বিত ভ্যাটের উপর সুদ (নোট ৩৫ এর ভিত্তিতে)', '41', { text: '0.00', alignment: 'right' }],
              ['বিলম্বিত সম্পূরক শুল্কের উপর সুদ (নোট ৩৭ এর ভিত্তিতে)', '42', { text: '0.00', alignment: 'right' }],
              ['রিটার্ন দাখিল না করার জন্য জরিমানা/দণ্ড', '43', { text: '0.00', alignment: 'right' }],
              ['অন্যান্য জরিমানা/দণ্ড/সুদ', '44', { text: '0.00', alignment: 'right' }],
              ['পরিশোধযোগ্য আবগারি শুল্ক', '45', { text: '0.00', alignment: 'right' }],
              ['পরিশোধযোগ্য উন্নয়ন অতিরিক্ত চার্জ', '46', { text: '0.00', alignment: 'right' }],
              ['পরিশোধযোগ্য আইসিটি উন্নয়ন অতিরিক্ত চার্জ', '47', { text: '0.00', alignment: 'right' }],
              ['পরিশোধযোগ্য স্বাস্থ্য অতিরিক্ত চার্জ', '48', { text: '0.00', alignment: 'right' }],
              ['পরিশোধযোগ্য পরিবেশ সমরক্ষা অতিরিক্ত চার্জ', '49', { text: '0.00', alignment: 'right' }],
              ['ট্রেজারিতে জমার জন্য নেট প্রদেয় ভ্যাট (35+41+43+44)', '50', { text: `(${Math.abs(n.note50 || 1979177.91).toLocaleString()})`, alignment: 'right', style: 'tBold' }],
              ['নেট প্রদেয় সম্পূরক শুল্ক কোষাগারে জমার জন্য (37+42)', '51', { text: '0.00', alignment: 'right' }],
              ['গত কর সময়কালের সমাপনী স্থিতি (ভ্যাট)', '52', { text: (n.note52 || 1445416.71).toLocaleString(), alignment: 'right', style: 'tBold' }],
              ['গত কর পর্বের সমাপনী স্থিতি (সম্পূরক শুল্ক)', '53', { text: '0.00', alignment: 'right' }]
            ]
          }
        },

        this.createFullWidthHeader("অংশ - ৮: পুরোনো অ্যাকাউন্টের বর্তমান ব্যালেন্সের জন্য সমন্বয়"),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['*', 40, 100], // Matches Section 7's stable 3-column layout
            body: [
              // Row 0: Header
              [
                { text: 'আইটেমসমূহ', style: 'tHead', alignment: 'center' },
                { text: 'নোট', style: 'tHead', alignment: 'center' },
                { text: 'পরিমাণ', style: 'tHead', alignment: 'center' }
              ],
              // Notes 54 - 57
              [
                'মূশক-১৮.৬ থেকে অবশিষ্ট ব্যালেন্স (ভ্যাট), [ বিধি ১১৮(৫)]',
                { text: '54', alignment: 'center' },
                { text: (n.note54 || '0.00').toLocaleString(), alignment: 'right' }
              ],
              [
                'মূশক-১৮.৬ থেকে অবশিষ্ট সম্পূরক শুল্ক (SD), [ বিধি ১১৮(৫)]',
                { text: '55', alignment: 'center' },
                { text: '0.00', alignment: 'right' }
              ],
              [
                'নোট ৫৪-এর জন্য হ্রাসমূলক সমন্বয় (নোট ৩৪-এর সর্বোচ্চ ৩০% পর্যন্ত)',
                { text: '56', alignment: 'center' },
                { text: '0.00', alignment: 'right' }
              ],
              [
                'নোট ৫৫-এর জন্য হ্রাসমূলক সমন্বয় (নোট ৩৬-এর সর্বোচ্চ ৩০% পর্যন্ত)',
                { text: '57', alignment: 'center' },
                { text: '0.00', alignment: 'right' }
              ]
            ]
          }
        },

        { text: '', pageBreak: 'before' },
        this.createFullWidthHeader("সেকশন - ৯: হিসাব কোড অনুযায়ী পরিশোধ সূচি (ট্রেজারি জমা)"),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['35%', '10%', '25%', '18%', '12%'],
            body: [
              [
                { text: 'আইটেমসমূহ', style: 'tHead', alignment: 'center' },
                { text: 'নোট', style: 'tHead', alignment: 'center' },
                { text: 'হিসাব কোড', style: 'tHead', alignment: 'center' },
                { text: 'পরিমাণ', style: 'tHead', alignment: 'center' },
                { text: '', style: 'tHead' }
              ],
              // Row 58: VAT Deposit
              ['বর্তমানের জন্য ভ্যাট জমা', '58', '1/1133/0030/0311', '0.00', 'সাবফর্ম'],
              // Row 59: SD Deposit
              ['বর্তমান কর সময়কালের জন্য সম্পূরক শুল্ক জমা', '59', '1/1133/0018/ 0711-0721', '0.00', 'সাবফর্ম'],
              // Row 60: Excise Duty
              ['আবগারি শুল্ক', '60', '1/1133/Acv‡ikbvj †KvW/0311', '0.00', 'সাবফর্ম'],
              // Row 61: Development Surcharge
              ['উন্নয়ন অতিরিক্ত চার্জ', '61', '1/1133/Acv‡ikbvj', '0.00', 'সাবফর্ম'],
              // Row 62: ICT Development Surcharge
              ['আইসিটি উন্নয়ন অতিরিক্ত চার্জ', '62', '1/1103/Acv‡ikbvj †KvW/1901', '0.00', 'সাবফর্ম'],
              // Row 63: Health Care Surcharge
              ['স্বাস্থ্যসেবা অতিরিক্ত চার্জ', '63', '1/1133/Acv‡ikbvj †KvW/0601', '0.00', 'সাবফর্ম'],
              // Row 64: Environmental Protection Surcharge
              ['পরিবেশ সুরক্ষা অতিরিক্ত শুল্ক', '64', '1/1103/Acv‡ikbvj †KvW/2225', '0.00', 'সাবফর্ম']
            ]
          }
        },

        this.createFullWidthHeader("সেকশন - ১০: সমাপনী ব্যালেন্স"),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['*', 40, 100], // Stable 3-column layout
            body: [
              // Row 0: Header
              [
                { text: 'আইটেমসমূহ', style: 'tHead', alignment: 'center' },
                { text: 'নোট', style: 'tHead', alignment: 'center' },
                { text: 'পরিমাণ', style: 'tHead', alignment: 'center' }
              ],
              // Row 65: Closing Balance (VAT)
              [
                'সমাপনী স্থিতি (ভ্যাট) [58 - (50 + 67) + অনুমোদিত নয় এমন ফেরতের পরিমাণ]',
                { text: '65', alignment: 'center' },
                { text: (n.note65 || 1979177.91).toLocaleString(), alignment: 'right', style: 'tBold' }
              ],
              // Row 66: Closing Balance (SD)
              [
                'সমাপনী স্থিতি (SD) [59 - (51 + 68) + অনুমোদিত নয় এমন ফেরতের পরিমাণ]',
                { text: '66', alignment: 'center' },
                { text: '0.00', alignment: 'right' }
              ]
            ]
          }
        },

        this.createFullWidthHeader("সেকশন - ১১: ফেরত"),
        {
          style: 'dataTable',
          table: {
            widths: ['35%', '35%', '10%', '20%'], // 4-column layout
            body: [
              // Header Row
              [
                { text: 'আমি আমার সমাপনী ব্যালেন্স ফেরত পেতে চাই', rowSpan: 3, margin: [0, 10] },
                { text: 'আইটেমসমূহ', style: 'tHead', alignment: 'center' },
                { text: 'নোট', style: 'tHead', alignment: 'center' },
                {
                  // columns: [
                  //   { text: 'হ্যাঁ', fontSize: 7 }, { text: '[  ]', fontSize: 7 },
                  //   { text: 'না', fontSize: 7 }, { text: '[  ]', fontSize: 7 }
                  // ],
                  columns: [
                    // Yes Option
                    { width: 'auto', table: { widths: [20], body: [[' ']] }, margin: [0, 0, 5, 0] },
                    { width: 'auto', text: 'হ্যাঁ', fontSize: 7, margin: [0, 2, 25, 0] },

                    // No Option
                    { width: 'auto', table: { widths: [20], body: [[' ']] }, margin: [0, 0, 5, 0] },
                    { width: 'auto', text: 'না', fontSize: 7, margin: [0, 2, 0, 0] }
                  ],
                  style: 'tHead'
                }
              ],
              // Note 67
              [
                {},
                'ফেরতের জন্য অনুরোধকৃত পরিমাণ (ভ্যাট)',
                { text: '67', alignment: 'center' },
                { text: (n.note67 || '0.00').toLocaleString(), alignment: 'right' }
              ],
              // Note 68
              [
                {},
                'ফেরতের জন্য অনুরোধকৃত পরিমাণ (সম্পূরক শুল্ক)',
                { text: '68', alignment: 'center' },
                { text: (n.note68 || '0.00').toLocaleString(), alignment: 'right' }
              ]
            ]
          }
        },

        this.createFullWidthHeader("সেকশন - ১২: ঘোষণা"),
        {
          style: 'dataTable',
          margin: [0, 0, 0, 0],
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: "আমি এই মর্মে ঘোষণা করছি যে এই রিটার্ন ফরমে প্রদত্ত সকল তথ্য সম্পূর্ণ, সত্য ও সঠিক। কোনো অসত্য/অসম্পূর্ণ বিবৃতির ক্ষেত্রে, আমি মূল্য সংযোজন কর ও সম্পূরক শুল্ক আইন, ২০১২ অথবা বর্তমানে প্রযোজ্য অন্য কোনো আইনের অধীনে দণ্ডনীয় ব্যবস্থার সম্মুখীন হতে পারি।",
                  fillColor: '#d9d9d9', // Gray background for the disclaimer
                  fontSize: 7,
                  margin: [5, 5, 5, 5]
                }
              ]
            ],
          }
        },
        {
          style: 'dataTable',
          table: {
            widths: ['35%', '5%', '60%'], // Matches the alignment of Section 1
            body: [
              ['নাম', ':', 'Hasanuzzaman'],
              ['পদবী', ':', ''],
              ['মোবাইল নম্বর', ':', ''],
              ['জাতীয় পরিচয় পত্র/পাসপোর্ট নম্বর', ':', ''],
              ['ইমেইল', ':', ''],
              ['স্বাক্ষর [ইলেকট্রনিক সাবমিশনের জন্য প্রয়োজন নেই]', ':', '']
            ]
          }
        }
      ],
      styles: {
        header: { font: 'BanglaFonts', fontSize: 8, bold: true, alignment: 'center' },
        subHeader: { font: 'BanglaFonts', fontSize: 7, bold: true, alignment: 'center', color: '#003366' },
        secHeaderCell: { font: 'BanglaFonts', fillColor: '#003366', color: 'white', bold: true, alignment: 'center', fontSize: 7, padding: [0, 2, 0, 2] },
        tHead: { font: 'BanglaFonts', fillColor: '#f2f2f2', bold: true, fontSize: 7 },
        tBold: { font: 'BanglaFonts', bold: true, fontSize: 7 },
        dataTable: { font: 'BanglaFonts', fontSize: 7, margin: [0, 0, 0, 5] },
        borderedTable: { font: 'BanglaFonts', margin: [0, 0, 0, 2] }
      }
    };
    pdfMake.createPdf(docDef).download('Mushak_9.1_Full_Report.pdf');
  }

}