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

  getMergedMushakData(apiEndpoint: string, lang: string): Observable<any> {
    const labels$ = this.http.get(`i18n/${lang}/dummyData.json`);
    const values$ = this.http.get(apiEndpoint);

    return forkJoin([labels$, values$]).pipe(
      map(([labels, values]: [any, any]) => {
        const mergedNotes: Record<string, any> = {};

        Object.keys(labels.notes).forEach(key => {
          mergedNotes[key] = values.notes?.[key] || { val: '0.00', sd: '0.00', vat: '0.00' };
        });

        return {
          labels: labels,
          notes: mergedNotes,
          taxpayer: values.taxpayer,
          returnSubmission: values.returnSubmission,
          mushak_4_3_data: values.mushak_values?.mushak_4_3_data || values.mushak_4_3_data || {}
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
  //   
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
      margin: [0, 0, 0, 0]
    };
  }

  exportFullMushakPdf(data: any, lang: string) {
    
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
      pageSize: 'A4',
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
            widths: ['35%', '2%', '63%'],
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
            widths: ['35%', '2%', '63%'],
            body: [
              [l.labels.tax_period, ':', { text: s.period || 'Oct / 2022', alignment: 'center' }],
              [l.labels.return_type, ':', {
                stack: [
                  { columns: [{ width: '70%', text: l.return_options ? l.return_options[0] : '' }, { table: { widths: ['30%'], body: [[' ']] }, margin: [0, 0, 10, 2], alignment: 'right' }] },
                  { columns: [{ width: '70%', text: l.return_options ? l.return_options[1] : '' }, { table: { widths: ['30%'], body: [[' ']] }, margin: [0, 0, 10, 2], alignment: 'right' }] },
                  { columns: [{ width: '70%', text: l.return_options ? l.return_options[2] : '' }, { table: { widths: ['30%'], body: [[' ']] }, margin: [0, 0, 10, 2], alignment: 'right' }] },
                  { columns: [{ width: '70%', text: l.return_options ? l.return_options[3] : '' }, { table: { widths: ['30%'], body: [[' ']] }, margin: [0, 0, 10, 2], alignment: 'right' }] }
                ], margin: [0, 2, 0, 1]
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
                  margin: [0, 2, 0, 1]
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
            widths: ['30%', '15%', '5%', '15%', '13%', '13%', '9%'],
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
              canvas: [{ type: 'rect', x: 0, y: 0, w: 535, h: 52, color: '#fcd5b4' }]
            },
            {
              text: l.labels.purchase_instruction.join('\n'),
              fontSize: 7,
              margin: [5, -50, 5, 2]
            }
          ],
          // margin: [0, 5, 0, 10]
        },
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['30%', '15%', '5%', '21%', '20%', '9%'],
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
              [{}, l.labels.import, '11', n.note11?.val || '0.00', { text: '', fillColor: '#d9d9d9' }, l.headers.sub_form],
              [{ text: l.notes.note11, rowSpan: 2 }, l.labels.local_purchase, '12', n.note12?.val || '0.00', { text: '', fillColor: '#d9d9d9' }, l.headers.sub_form],
              [{}, l.labels.import, '13', n.note13?.val || '0.00', { text: '', fillColor: '#d9d9d9' }, l.headers.sub_form],

              // Standard Rated - Main Data (Notes 14-15)
              [{ text: l.notes.note12, rowSpan: 2 }, l.labels.local_purchase, '14', n.note14?.val || '0.00', n.note14?.vat || '0.00', l.headers.sub_form],
              [{}, l.labels.import, '15', n.note15?.val || '0.00', n.note15?.vat || '0.00', l.headers.sub_form],

              // Other Categories (Notes 16-22)
              [{ text: l.notes.note13, rowSpan: 2 }, l.labels.local_purchase, '16', n.note16?.val || '0.00', n.note16?.vat || '0.00', l.headers.sub_form],
              [{}, l.labels.import, '17', n.note17?.val || '0.00', n.note17?.vat || '0.00', l.headers.sub_form],
              [{ text: l.notes.note14, rowSpan: 1 }, l.labels.local_purchase, '18', n.note18?.val || '0.00', n.note18?.vat || '0.00', l.headers.sub_form],
              [{ text: l.notes.note15, rowSpan: 2 }, l.labels.from_turnover, '19', n.note19?.val || '0.00', n.note19?.vat || '0.00', l.headers.sub_form],
              [{}, l.labels.from_unregistered, '20', n.note20?.val || '0.00', n.note20?.vat || '0.00', l.headers.sub_form],
              [{ text: l.notes.note16, rowSpan: 2 }, l.labels.local_purchase, '21', n.note21?.val || '0.00', n.note21?.vat || '0.00', l.headers.sub_form],
              [{}, l.labels.import, '22', n.note22?.val || n.note22?.val || '0.00', n.note22?.vat || '0.00', l.headers.sub_form],

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
            widths: ['45%', '5%', '41%', '9%'],
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
            widths: ['45%', '5%', '41%', '9%'],
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
            widths: ['70%', '5%', '25%'],
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
              [l.notes.note50, '50', { text: formatAmount(n.note50?.val || n.note50), alignment: 'right' }],
              [l.notes.note51, '51', { text: formatAmount(n.note51?.val || n.note51), alignment: 'right' }],
              [l.notes.note52, '52', { text: formatAmount(n.note52?.val || n.note52), alignment: 'right' }],
              [l.notes.note53, '53', { text: formatAmount(n.note53?.val || n.note53), alignment: 'right' }]
            ]
          }
        },

        this.createFullWidthHeader(l.sections.s8),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['70%', '5%', '25%'],
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
            widths: ['43%', '5%', '25%', '18%', '9%'],
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
            widths: ['65%', '5%', '30%'],
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
            widths: ['35%', '35%', '5%', '25%'],
            body: [
              // Header Row
              [
                { text: l.labels.interest_refund, rowSpan: 3, margin: [0, 10] },
                { text: l.headers.items, style: 'tHead', alignment: 'center' },
                { text: l.headers.note, style: 'tHead', alignment: 'center' },
                {
                  columns: [
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
                  fillColor: '#d9d9d9',
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
            widths: ['38%', '2%', '60%'],
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


  // --- MERGED MUSHAK-9.1 EXCEL (ALL SECTIONS 1-12) ---
  async exportFullMushakExcel(data: any, lang: string) {
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


  exportFullMushakPdfBangla(data: any, lang: string) {
    const l = data.labels || {};
    const n = data?.notes || {};
    const t = data?.taxpayer || {};
    const s = data?.returnSubmission || {};

    (pdfMake as any).fonts = {
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/kalpurush.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf'
      }
    };

    const formatAmount = (val: any) => {
      const num = parseFloat(val) || 0;
      return num < 0 ? `(${Math.abs(num)})` : num.toFixed(2);
    };

    const docDef: any = {
      pageSize: 'A4',
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
            widths: ['35%', '2%', '63%'],
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
            widths: ['35%', '2%', '63%'],
            body: [
              [l.labels.tax_period, ':', { text: s.period || 'Oct / 2022', alignment: 'center' }],
              [l.labels.return_type, ':', {
                stack: [
                  { columns: [{ width: '70%', text: l.return_options ? l.return_options[0] : '' }, { table: { widths: ['30%'], body: [[' ']] }, margin: [0, 0, 10, 2], alignment: 'right' }] },
                  { columns: [{ width: '70%', text: l.return_options ? l.return_options[1] : '' }, { table: { widths: ['30%'], body: [[' ']] }, margin: [0, 0, 10, 2], alignment: 'right' }] },
                  { columns: [{ width: '70%', text: l.return_options ? l.return_options[2] : '' }, { table: { widths: ['30%'], body: [[' ']] }, margin: [0, 0, 10, 2], alignment: 'right' }] },
                  { columns: [{ width: '70%', text: l.return_options ? l.return_options[3] : '' }, { table: { widths: ['30%'], body: [[' ']] }, margin: [0, 0, 10, 2], alignment: 'right' }] }
                ], margin: [0, 2, 0, 1]
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
                  margin: [0, 2, 0, 1]
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
            widths: ['30%', '15%', '5%', '15%', '13%', '13%', '9%'],
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
                l.notes.note1.split('-')[1] || l.notes.note1, '১',
                (n['note1']?.val || 0).toLocaleString(),
                (n['note1']?.sd || 0).toLocaleString(),
                (n['note1']?.vat || 0).toLocaleString(),
                l.headers.sub_form
              ],
              [
                {}, l.notes.note2.split('-')[1] || l.notes.note2, '২',
                (n['note2']?.val || 0).toLocaleString(),
                (n['note2']?.sd || 0).toLocaleString(),
                (n['note2']?.vat || 0).toLocaleString(),
                l.headers.sub_form
              ],
              // Note 3: Exempted
              [{ text: l.notes.note3, colSpan: 2 }, {}, '৩', (n['note3']?.val || 0).toLocaleString(), '0.00', '0.00', l.headers.sub_form],
              // Note 4: Standard Rated
              [
                { text: l.notes.note4, colSpan: 2 }, {}, '৪',
                (n['note4']?.val || 0).toLocaleString(undefined, { minimumFractionDigits: 2 }),
                (n['note4']?.sd || 0).toLocaleString(),
                (n['note4']?.vat || 0).toLocaleString(undefined, { minimumFractionDigits: 2 }),
                l.headers.sub_form
              ],
              // Note 5-8: Other Categories
              [{ text: l.notes.note5, colSpan: 2 }, {}, '৫', (n['note5']?.val || 0).toLocaleString(), '0.00', '0.00', l.headers.sub_form],
              [{ text: l.notes.note6, colSpan: 2 }, {}, '৬', (n['note6']?.val || 0).toLocaleString(), '0.00', '0.00', l.headers.sub_form],
              [{ text: l.notes.note7, colSpan: 2 }, {}, '৭', (n['note7']?.val || 0).toLocaleString(), '0.00', '0.00', l.headers.sub_form],
              [{ text: l.notes.note8, colSpan: 2 }, {}, '৮', (n['note8']?.val || 0).toLocaleString(), '0.00', '0.00', l.headers.sub_form],
              // Note 9: Total
              [
                { text: l.notes.note9, colSpan: 2, style: 'tBold' }, {},
                { text: '৯', style: 'tBold' },
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
              canvas: [{ type: 'rect', x: 0, y: 0, w: 535, h: 42, color: '#fcd5b4' }]
            },
            {
              text: l.labels.purchase_instruction.join('\n'),
              fontSize: 7,
              margin: [5, -38, 5, 2]
            }
          ],
          // margin: [0, 5, 0, 10]
        },
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['30%', '15%', '5%', '21%', '20%', '9%'],
            body: [
              [
                { text: l.headers.nature_purchase, style: 'tHead', colSpan: 2, alignment: 'center' },
                {},
                { text: l.headers.note, style: 'tHead', alignment: 'center' },
                { text: l.headers.value, style: 'tHead', alignment: 'center' },
                { text: l.headers.vat, style: 'tHead', alignment: 'center' },
                { text: '', border: [false, false, false, false] }
              ],
              // Zero Rated & Exempted (Notes 10-13)
              [{ text: l.notes.note10, rowSpan: 2 }, l.labels.local_purchase, '10', n.note10?.val || '0.00', { text: '', fillColor: '#d9d9d9' }, l.headers.sub_form],
              [{}, l.labels.import, '11', n.note11?.val || '0.00', { text: '', fillColor: '#d9d9d9' }, l.headers.sub_form],
              [{ text: l.notes.note11, rowSpan: 2 }, l.labels.local_purchase, '12', n.note12?.val || '0.00', { text: '', fillColor: '#d9d9d9' }, l.headers.sub_form],
              [{}, l.labels.import, '13', n.note13?.val || '0.00', { text: '', fillColor: '#d9d9d9' }, l.headers.sub_form],

              // Standard Rated - Main Data (Notes 14-15)
              [{ text: l.notes.note12, rowSpan: 2 }, l.labels.local_purchase, '14', n.note14?.val || '0.00', n.note14?.vat || '0.00', l.headers.sub_form],
              [{}, l.labels.import, '15', n.note15?.val || '0.00', n.note15?.vat || '0.00', l.headers.sub_form],

              // Other Categories (Notes 16-22)
              [{ text: l.notes.note13, rowSpan: 2 }, l.labels.local_purchase, '16', n.note16?.val || '0.00', n.note16?.vat || '0.00', l.headers.sub_form],
              [{}, l.labels.import, '17', n.note17?.val || '0.00', n.note17?.vat || '0.00', l.headers.sub_form],
              [{ text: l.notes.note14, rowSpan: 1 }, l.labels.local_purchase, '18', n.note18?.val || '0.00', n.note18?.vat || '0.00', l.headers.sub_form],
              [{ text: l.notes.note15, rowSpan: 2 }, l.labels.from_turnover, '19', n.note19?.val || '0.00', n.note19?.vat || '0.00', l.headers.sub_form],
              [{}, l.labels.from_unregistered, '20', n.note20?.val || '0.00', n.note20?.vat || '0.00', l.headers.sub_form],
              [{ text: l.notes.note16, rowSpan: 2 }, l.labels.local_purchase, '21', n.note21?.val || '0.00', n.note21?.vat || '0.00', l.headers.sub_form],
              [{}, l.labels.import, '22', n.note22?.val || n.note22?.val || '0.00', n.note22?.vat || '0.00', l.headers.sub_form],

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
            widths: ['45%', '5%', '41%', '9%'],
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
            widths: ['45%', '5%', '41%', '9%'],
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
            widths: ['70%', '5%', '25%'],
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
              [l.notes.note50, '50', { text: formatAmount(n.note50?.val || n.note50), alignment: 'right' }],
              [l.notes.note51, '51', { text: formatAmount(n.note51?.val || n.note51), alignment: 'right' }],
              [l.notes.note52, '52', { text: formatAmount(n.note52?.val || n.note52), alignment: 'right' }],
              [l.notes.note53, '53', { text: formatAmount(n.note53?.val || n.note53), alignment: 'right' }]
            ]
          }
        },

        this.createFullWidthHeader(l.sections.s8),
        {
          style: 'dataTable',
          table: {
            headerRows: 1,
            widths: ['70%', '5%', '25%'],
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
            widths: ['43%', '5%', '25%', '18%', '9%'],
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
            widths: ['65%', '5%', '30%'],
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
            widths: ['35%', '35%', '5%', '25%'],
            body: [
              // Header Row
              [
                { text: l.labels.interest_refund, rowSpan: 3, margin: [0, 10] },
                { text: l.headers.items, style: 'tHead', alignment: 'center' },
                { text: l.headers.note, style: 'tHead', alignment: 'center' },
                {
                  columns: [
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
                  fillColor: '#d9d9d9',
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
            widths: ['38%', '2%', '60%'],
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

  exportInputOutputCoefficientEnglish(data: any, lang: string) {
    const l = (data.labels.mushak_4_3 || {}) as any;
    const f = (l.footer || {}) as any;

    // Data mapping from mushak_values
    const mainData = data.mushak_values?.mushak_4_3_data[lang] || data.mushak_4_3_data[lang] || {};
    const info = (mainData.companyInfo || {}) as any;
    const items = (mainData.items || []) as any[];

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-Regular.ttf'
      }
    };

    const safe = (val: any) => val !== undefined && val !== null ? val.toString() : '';

    const docDef: any = {
      pageSize: 'A4',
      pageOrientation: 'landscape',
      defaultStyle: { font: 'Nunito', fontSize: 8 },
      content: [
        // Header Section
        {
          columns: [
            { text: '', width: '*' },
            {
              stack: [
                { text: l.titles.gov, style: 'header' },
                { text: l.titles.nbr, style: 'header' },
                { text: l.titles.form, style: 'subHeader' },
                { text: l.titles.rule, style: 'subHeader' }
              ], width: 400
            },
            { text: l.titles.m_name, alignment: 'right', bold: true, fontSize: 12, width: '*' }
          ]
        },

        // Institution Information
        {
          margin: [0, 15, 0, 10],
          table: {
            widths: ['30%', '2%', '68%'],
            body: [
              [l.info.comp_name, ':', info.name],
              [l.info.address, ':', info.address],
              [l.info.bin, ':', safe(info.bin)],
              [l.info.sub_date, ':', safe(info.submissionDate)],
              [l.info.first_supply, ':', safe(info.firstSupplyDate)]
            ]
          },
          layout: 'noBorders'
        },

        // Main Data Table (12 Columns)
        {
          table: {
            headerRows: 2,
            widths: [25, 55, 90, 60, 90, 45, 45, 45, 40, 80, 50, 45],
            body: [
              // Row 1: Merged Headers 
              [
                { text: l.headers.sl, rowSpan: 2, alignment: 'center', bold: true },
                { text: l.headers.hs_code, rowSpan: 2, alignment: 'center', bold: true },
                { text: l.headers.item_desc, rowSpan: 2, alignment: 'center', bold: true },
                { text: l.headers.unit, rowSpan: 2, alignment: 'center', bold: true },
                { text: 'Description of Raw Materials, Quantity & Purchase Price', colSpan: 5, alignment: 'center', bold: true },
                {}, {}, {}, {},
                { text: 'Value Addition Details', colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: l.headers.remarks, rowSpan: 2, alignment: 'center', bold: true }
              ],
              // Row 2: Sub-headers
              [
                {}, {}, {}, {},
                { text: l.headers.raw_material, bold: true }, { text: l.headers.buy_price, bold: true },
                { text: l.headers.qty_w, bold: true }, { text: l.headers.qty_wo, bold: true },
                { text: l.headers.wastage_p, bold: true }, { text: l.headers.va_sector, bold: true },
                { text: l.headers.va_value, bold: true }, {}
              ],
              // Data Mapping
              ...items.map((item, idx) => [
                { text: (idx + 1).toString(), alignment: 'center' },
                safe(item.hsCode),
                safe(item.itemName),
                safe(item.unit),
                safe(item.rawMaterialName),
                { text: safe(item.price), alignment: 'right' },
                { text: safe(item.qtyInclWastage), alignment: 'right' },
                { text: safe(item.wastageQty), alignment: 'right' },
                { text: safe(item.wastagePercent) + '%', alignment: 'right' },
                safe(item.vaSector),
                { text: safe(item.vaValue), alignment: 'right' },
                safe(item.remarks)
              ])
            ]
          }
        },

        // Footer Section [cite: 30, 31, 32, 33]
        {
          margin: [0, 20, 0, 10],
          columns: [
            { text: '', width: '*' },
            {
              stack: [
                { text: f.auth_person_title, bold: true },
                { text: f.designation, margin: [0, 5, 0, 5] },
                { text: f.signature },
                { text: f.seal, margin: [0, 5, 0, 0] }
              ],
              width: 280
            }
          ]
        },
        {
          stack: [
            { text: f.special_note_title, bold: true, decoration: 'underline', margin: [0, 10, 0, 5] },
            {
              ol: f.notes || [],
              fontSize: 7,
              lineHeight: 1.3
            }
          ]
        }
      ],
      styles: {
        header: { fontSize: 11, bold: true, alignment: 'center' },
        subHeader: { fontSize: 9, alignment: 'center' }
      }
    };

    pdfMake.createPdf(docDef).download('mushak_4_3_English_Report.pdf');
  }

  exportInputOutputCoefficientBangla(data: any, lang: string) {
    const l = (data.labels.mushak_4_3 || {}) as any;
    const info = data.mushak_4_3_data[lang]?.companyInfo || {};
    const items = (data.mushak_4_3_data[lang]?.items || []) as any[];
    const f = l.footer || {};

    (pdfMake as any).fonts = {
      kalpurush: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/kalpurush.ttf',
      }
    };

    const safeText = (val: any, isNum = false) => {
      return val !== undefined && val !== null ? val.toString() : '';
    };

    const docDef: any = {
      pageSize: 'A4',
      pageOrientation: 'landscape',
      defaultStyle: { font: 'kalpurush', fontSize: 8 },
      content: [
        {
          columns: [
            { text: '', width: '*' },
            {
              stack: [
                { text: l.titles.gov, style: 'header' },
                { text: l.titles.nbr, style: 'header' },
                { text: l.titles.form, style: 'subHeader' },
                { text: l.titles.rule, style: 'subHeader' }
              ], width: 300
            },
            { text: l.titles.m_name, alignment: 'right', bold: true, fontSize: 12, width: '*' }
          ]
        },

        {
          margin: [0, 15, 0, 10],
          table: {
            widths: ['25%', '2%', '73%'],
            body: [
              [l.info.comp_name, ':', info.name],
              [l.info.address, ':', info.address],
              [l.info.bin, ':', safeText(info.bin, true)],
              [l.info.sub_date, ':', safeText(info.submissionDate, true)],
              [l.info.first_supply, ':', safeText(info.firstSupplyDate, true)]
            ]
          },
          layout: 'noBorders'
        },

        // মূল টেবিল (১২টি কলাম)
        {
          table: {
            headerRows: 2,
            widths: [25, 50, 80, 50, 80, 50, 45, 45, 35, 75, 50, 40],
            body: [
              [
                { text: l.headers.sl, rowSpan: 2, alignment: 'center' },
                { text: l.headers.hs_code, rowSpan: 2, alignment: 'center' },
                { text: l.headers.item_desc, rowSpan: 2, alignment: 'center' },
                { text: l.headers.unit, rowSpan: 2, alignment: 'center' },
                { text: l.headers.item_desc_wastage_pencentage, colSpan: 5, alignment: 'center' },
                {}, {}, {}, {},
                { text: l.headers.price_correction, colSpan: 2, alignment: 'center' },
                {},
                { text: l.headers.remarks, rowSpan: 2, alignment: 'center' }
              ],
              [
                {}, {}, {}, {},
                l.headers.raw_material, l.headers.buy_price, l.headers.qty_w, l.headers.qty_wo, l.headers.wastage_p,
                l.headers.va_sector, l.headers.va_value, {}
              ],
              ...items.map((item, idx) => [
                { text: (idx + 1).toString(), alignment: 'center' },
                safeText(item.hsCode, true),
                safeText(item.itemName),
                safeText(item.unit),
                safeText(item.rawMaterialName),
                safeText(item.price, true),
                safeText(item.qtyInclWastage, true),
                safeText(item.wastageQty, true),
                { text: item.wastagePercent + '%', alignment: 'right' },
                safeText(item.vaSector),
                safeText(item.vaValue, true),
                safeText(item.remarks)
              ])
            ]
          }
        },
        {
          margin: [0, 20, 0, 10],
          columns: [
            { text: '', width: '*' },
            {
              stack: [
                { text: f.auth_person_title, bold: true },
                { text: f.designation, margin: [0, 5, 0, 5] },
                { text: f.signature },
                { text: f.seal, margin: [0, 5, 0, 0] }
              ],
              width: 250,
              alignment: 'left'
            }
          ]
        },

        // ৩. বিশেষ দ্রষ্টব্য সেকশন (বামে)
        {
          stack: [
            { text: f.special_note_title, bold: true, decoration: 'underline', margin: [0, 10, 0, 5] },
            {
              text: (f.notes || []).join('\n'),
              fontSize: 7,
              lineHeight: 1.4
            }
          ]
        }
      ],
      styles: {
        header: { fontSize: 11, bold: true, alignment: 'center' },
        subHeader: { fontSize: 9, alignment: 'center' }
      }
    };

    pdfMake.createPdf(docDef).download('mushak_4_3_Report.pdf');
  }

  exportmushak_6_1English(data: any, lang: string) {
    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-Regular.ttf'
      }
    };

    const l = (data.labels?.mushak_6_1 || {}) as any;
    const targetData = data.mushak_values?.[lang] || data[lang] || {};
    const m61 = (targetData.mushak_6_1_data || {}) as any;
    const info = (m61.companyInfo || {}) as any;
    const items = (m61.items || []) as any[];

    const safe = (val: any) => (val !== undefined && val !== null) ? val.toString() : ' ';

    const docDef: any = {
      pageSize: 'A4',
      pageOrientation: 'landscape',
      defaultStyle: { font: 'Nunito', fontSize: 6 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        { text: safe(l.titles?.form), alignment: 'center', bold: true, fontSize: 10 },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 7, margin: [0, 0, 0, 5] },
        { text: safe(l.titles?.sub_title), alignment: 'center', decoration: 'underline' },

        // Institution Information Section
        {
          margin: [0, 10, 0, 10],
          table: {
            widths: ['25%', '2%', '73%'],
            body: [
              [l.info?.comp_name, ':', safe(info.name)],
              [l.info?.address, ':', safe(info.address)],
              [l.info?.bin, ':', safe(info.bin)]
            ]
          },
          layout: 'noBorders'
        },

        // Main Table (21 Columns as per PDF)
        {
          table: {
            headerRows: 3,
            widths: [15, 30, 25, 25, 30, 30, 35, 35, 35, 40, 30, 30, 25, 25, 30, 30, 30, 30, 30, 30, 30],
            body: [
              // Row 1: Merged Headers
              [
                { text: l.headers?.sl, rowSpan: 2, bold: true, alignment: 'center' },
                { text: l.headers?.date, rowSpan: 2, bold: true, alignment: 'center' },
                { text: l.headers?.opening_stock, colSpan: 2, alignment: 'center', bold: true }, {},
                { text: l.headers?.invoice_info, colSpan: 2, alignment: 'center', bold: true }, {},
                { text: l.headers?.seller_info, colSpan: 3, alignment: 'center', bold: true }, {}, {},
                { text: l.headers?.item_desc, rowSpan: 2, bold: true, alignment: 'center' },
                { text: l.headers?.purchase_info, colSpan: 4, alignment: 'center', bold: true }, {}, {}, {},
                { text: l.headers?.total_materials, colSpan: 2, alignment: 'center', bold: true }, {},
                { text: l.headers?.usage_info, colSpan: 2, alignment: 'center', bold: true }, {},
                { text: l.headers?.closing_stock, colSpan: 2, alignment: 'center', bold: true }, {},
                { text: l.headers?.remarks, rowSpan: 2, bold: true, alignment: 'center' }
              ],
              // Row 2: Sub-headers
              [
                {}, {}, 'Qty', 'Value', 'No', 'Date', 'Name', 'Address', 'BIN', '', 'Qty', 'Value', 'SD', 'VAT', 'Qty', 'Value', 'Qty', 'Value', 'Qty', 'Value', ''
              ],
              // Row 3: Column Reference Numbers (1-21)
              Array.from({ length: 21 }, (_, i) => ({ text: `(${i + 1})`, alignment: 'center', fontSize: 5 })),

              // Data Mapping from items
              ...items.map((item: any) => [
                safe(item.sl), safe(item.date), safe(item.opening_qty), safe(item.opening_val),
                safe(item.invoice_no), safe(item.invoice_date), safe(item.seller_name),
                safe(item.seller_address), safe(item.seller_bin), safe(item.item_desc),
                safe(item.purchase_qty), safe(item.purchase_val), safe(item.sd), safe(item.vat),
                safe(item.total_qty), safe(item.total_val), safe(item.used_qty), safe(item.used_val),
                safe(item.closing_qty), safe(item.closing_val), safe(item.remarks)
              ])
            ]
          }
        }
      ]
    };
    pdfMake.createPdf(docDef).download(`mushak_6_1_English.pdf`);
  }

  exportmushak_6_1Bangla(data: any, lang: string) {
    
    const l = (data.labels?.mushak_6_1 || {}) as any;
    const targetData = data.labels?.mushak_6_1 || {}; 
    const info = (targetData.info || {}) as any;
    const items = (targetData.items || []) as any[];
    const sh = (data.labels?.mushak_6_1?.sub_headers || {}) as any;

    const safe = (val: any) => val !== undefined && val !== null ? val.toString() : ' ';

    (pdfMake as any).fonts = {
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/kalpurush.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf'
      }
    };

    const docDef: any = {
      pageSize: 'A4',
      pageOrientation: 'landscape',
      defaultStyle: { font: 'PlaywriteCU', fontSize: 6 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },  
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        { text: safe(l.titles?.form), alignment: 'center', bold: true, fontSize: 10 },  
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 7 },  

        {
          margin: [0, 10, 0, 10],
          table: {
            widths: ['25%', '2%', '73%'],
            body: [
              [l.info?.comp_name, ':', safe(info.name)],
              [l.info?.address, ':', safe(info.address)],
              [l.info?.bin, ':', safe(info.bin)]
            ]
          },
          layout: 'noBorders'
        },

        {
          table: {
            headerRows: 3,
            widths: [15, 30, 25, 25, 30, 30, 35, 35, 35, 40, 30, 30, 25, 25, 30, 30, 30, 30, 30, 30, 30],
            body: [
              // Row 1: Merged Headers 
              [
                { text: l.headers?.sl, rowSpan: 2, bold: true },
                { text: l.headers?.date, rowSpan: 2, bold: true },
                { text: l.headers?.opening_stock, colSpan: 2, alignment: 'center', bold: true }, {},
                { text: l.headers?.invoice_info, colSpan: 2, alignment: 'center', bold: true }, {},
                { text: l.headers?.seller_info, colSpan: 3, alignment: 'center', bold: true }, {}, {},
                { text: l.headers?.item_desc, rowSpan: 2, bold: true },
                { text: l.headers?.purchase_info, colSpan: 4, alignment: 'center', bold: true }, {}, {}, {},
                { text: l.headers?.total_materials, colSpan: 2, alignment: 'center', bold: true }, {},
                { text: l.headers?.usage_info, colSpan: 2, alignment: 'center', bold: true }, {},
                { text: l.headers?.closing_stock, colSpan: 2, alignment: 'center', bold: true }, {},
                { text: l.headers?.remarks, rowSpan: 2, bold: true }
              ],
              // Row 2: Sub-headers 
              [
                {}, {},
                sh.qty || ' ', sh.val || ' ', // opening
                sh.no || ' ', sh.date || ' ', // invoice
                sh.name || ' ', sh.addr || ' ', sh.bin || ' ', // seller
                {},
                sh.qty || ' ', sh.val || ' ', sh.sd || ' ', sh.vat || ' ', // purchase
                sh.qty || ' ', sh.val || ' ', // total
                sh.qty || ' ', sh.val || ' ', // usage
                sh.qty || ' ', sh.val || ' ', // closing
                {}
              ],
              // Row 3: Column Numbers (1) to (21) 
              Array.from({ length: 21 }, (_, i) => ({ text: `(${i + 1})`, alignment: 'center', fontSize: 5 })),

              // Data Rows from db.json
              ...items.map((item: any) => [
                safe(item.sl), safe(item.date), safe(item.opening_qty), safe(item.opening_val),
                safe(item.invoice_no), safe(item.invoice_date), safe(item.seller_name),
                safe(item.seller_address), safe(item.seller_bin), safe(item.item_desc),
                safe(item.purchase_qty), safe(item.purchase_val), safe(item.sd), safe(item.vat),
                safe(item.total_qty), safe(item.total_val), safe(item.used_qty), safe(item.used_val),
                safe(item.closing_qty), safe(item.closing_val), safe(item.remarks)
              ])
            ]
          }
        },

        { text: '\n' + safe(l.footer?.note_title), bold: true, decoration: 'underline' },
        {
          ol: l.footer?.notes || [],
          fontSize: 6.5
        }
      ]
    };
    pdfMake.createPdf(docDef).download(`mushak_6_1_${lang}.pdf`);
  }
}