import { Injectable } from '@angular/core';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

// pdfmake imports
import pdfMakeImport from 'pdfmake/build/pdfmake';
import pdfFontsImport from 'pdfmake/build/vfs_fonts';
import { MushakData } from './MushakData';
import { forkJoin, map, Observable } from 'rxjs';
import { HttpClient } from '@angular/common/http';

//QR code generation
import QRCode from 'qrcode';

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

        Object.keys(labels.notes).forEach((key) => {
          mergedNotes[key] = values.notes?.[key] || { val: '0.00', sd: '0.00', vat: '0.00' };
        });

        return {
          labels: labels,
          notes: mergedNotes,
          taxpayer: values.taxpayer,
          returnSubmission: values.returnSubmission,
          mushak_2_3_data: values.mushak_values?.mushak_2_3_data || values.mushak_2_3_data || {},
          mushak_2_1_data: values.mushak_values?.mushak_2_1_data || values.mushak_2_1_data || {},
          mushak_4_3_data: values.mushak_values?.mushak_4_3_data || values.mushak_4_3_data || {},
          mushak_6_1_data: values.mushak_values?.mushak_6_1_data || values.mushak_6_1_data || {},
          mushak_6_2_data: values.mushak_values?.mushak_6_2_data || values.mushak_6_2_data || {},
          mushak_6_2_1_data:
            values.mushak_values?.mushak_6_2_1_data || values.mushak_6_2_1_data || {},
          mushak_6_3_data: values.mushak_values?.mushak_6_3_data || values.mushak_6_3_data || {},
          mushak_6_4_data: values.mushak_values?.mushak_6_4_data ||
            values.mushak_6_4_data || { items: [] },
          mushak_6_5_data: values.mushak_values?.mushak_6_5_data ||
            values.mushak_6_5_data || { items: [] },
          mushak_6_6_data: values.mushak_values?.mushak_6_6_data ||
            values.mushak_6_6_data || { items: [] },
          mushak_6_7_data: values.mushak_values?.mushak_6_7_data ||
            values.mushak_6_7_data || { items: [] },
          mushak_6_8_data: values.mushak_values?.mushak_6_8_data ||
            values.mushak_6_8_data || { items: [] },
          mushak_6_9_data: values.mushak_values?.mushak_6_9_data ||
            values.mushak_6_9_data || { items: [] },
          mushak_6_10_data: values.mushak_values?.mushak_6_10_data ||
            values.mushak_6_10_data || { items: [] },
          mushak_6_11_data: values.mushak_values?.mushak_6_11_data ||
            values.mushak_6_11_data || { items: [] },
          mushak_10_1_data: values.mushak_values?.mushak_10_1_data ||
            values.mushak_10_1_data || { items: [] },
          mushak_18_1_data: values.mushak_values?.mushak_18_1_data ||
            values.mushak_18_1_data || { items: [] },
          mushak_18_2_data: values.mushak_values?.mushak_18_2_data || values.mushak_18_2_data || {},
          mushak_18_3_data: values.mushak_values?.mushak_18_3_data || values.mushak_18_3_data || {},
        };
      }),
    );
  }

  private mergeNotes(labelNotes: any, valueNotes: any): any {
    const merged: any = {};
    Object.keys(labelNotes).forEach((key) => {
      merged[key] = {
        label: labelNotes[key],
        val: valueNotes?.[key]?.val ?? 0,
        vat: valueNotes?.[key]?.vat ?? 0,
        sd: valueNotes?.[key]?.sd ?? 0,
      };
    });
    return merged;
  }

  // getMushakJsonData(): Observable<MushakData> {
  //
  //   return this.http.get<MushakData>('i18n/EN/dummyData.json');
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
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
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
        String(r.HSCode ?? ''),
        String(r.Description ?? ''),
        String(r.CD ?? 0),
        String(r.SD ?? 0),
        String(r.VAT ?? 0),
        String(r.AIT ?? 0),
        String(r.RD ?? 0),
        String(r.TTI ?? 0),
      ]),
    ];

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [20, 30, 20, 30],
      content: [
        { text: 'VAT HS Code', style: 'title' },
        {
          table: { headerRows: 1, widths: [60, '*', 30, 30, 35, 35, 30, 30], body },
          layout: 'lightHorizontalLines',
        },
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
        body: [[{ text: text, style: 'secHeaderCell' }]],
      },
      layout: 'noBorders',
      margin: [0, 0, 0, 0],
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
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
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
        fontSize: 7,
      },
      content: [
        {
          stack: [
            { text: l.titles.gov, style: 'header' },
            { text: l.titles.nbr, style: 'header' },
            { text: `\n${l.titles.form}`, style: 'subHeader' },
            { text: `${l.titles.rule}\n`, style: 'subHeader' },
            { text: '\n', style: 'subHeader' },
          ],
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
              [l.labels.activity, ':', t.activity],
            ],
          },
        },

        this.createFullWidthHeader(l.sections.s2),
        {
          style: 'dataTable',
          table: {
            widths: ['35%', '2%', '63%'],
            body: [
              [l.labels.tax_period, ':', { text: s.period || 'Oct / 2022', alignment: 'center' }],
              [
                l.labels.return_type,
                ':',
                {
                  stack: [
                    {
                      columns: [
                        { width: '70%', text: l.return_options ? l.return_options[0] : '' },
                        {
                          table: { widths: ['30%'], body: [[' ']] },
                          margin: [0, 0, 10, 2],
                          alignment: 'right',
                        },
                      ],
                    },
                    {
                      columns: [
                        { width: '70%', text: l.return_options ? l.return_options[1] : '' },
                        {
                          table: { widths: ['30%'], body: [[' ']] },
                          margin: [0, 0, 10, 2],
                          alignment: 'right',
                        },
                      ],
                    },
                    {
                      columns: [
                        { width: '70%', text: l.return_options ? l.return_options[2] : '' },
                        {
                          table: { widths: ['30%'], body: [[' ']] },
                          margin: [0, 0, 10, 2],
                          alignment: 'right',
                        },
                      ],
                    },
                    {
                      columns: [
                        { width: '70%', text: l.return_options ? l.return_options[3] : '' },
                        {
                          table: { widths: ['30%'], body: [[' ']] },
                          margin: [0, 0, 10, 2],
                          alignment: 'right',
                        },
                      ],
                    },
                  ],
                  margin: [0, 2, 0, 1],
                },
              ],
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
                            {
                              width: 'auto',
                              table: { widths: [20], body: [[' ']] },
                              margin: [0, 0, 5, 0],
                            },
                            {
                              width: 'auto',
                              text: l.labels.yes,
                              fontSize: 7,
                              margin: [0, 2, 25, 0],
                            },

                            // No Option
                            {
                              width: 'auto',
                              table: { widths: [20], body: [[' ']] },
                              margin: [0, 0, 5, 0],
                            },
                            { width: 'auto', text: l.labels.no, fontSize: 7, margin: [0, 2, 0, 0] },
                          ],
                        },
                        { width: '*', text: '' },
                      ],
                    },
                    {
                      text: l.labels.activity_note,
                      fontSize: 7,
                      alignment: 'center',
                      margin: [0, 5, 0, 0],
                      color: '#333333',
                    },
                  ],
                  margin: [0, 2, 0, 1],
                },
              ],
              [l.labels.sub_date, ':', { text: s.date || 'Oct / 2022', alignment: 'center' }],
            ],
          },
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
                { text: '', border: [false, false, false, false] },
              ],
              // Note 1 & 2: Zero Rated
              [
                { text: l.notes.note1.split('-')[0], rowSpan: 2 },
                l.notes.note1.split('-')[1] || l.notes.note1,
                '1',
                (n['note1']?.val || 0).toLocaleString(),
                (n['note1']?.sd || 0).toLocaleString(),
                (n['note1']?.vat || 0).toLocaleString(),
                l.headers.sub_form,
              ],
              [
                {},
                l.notes.note2.split('-')[1] || l.notes.note2,
                '2',
                (n['note2']?.val || 0).toLocaleString(),
                (n['note2']?.sd || 0).toLocaleString(),
                (n['note2']?.vat || 0).toLocaleString(),
                l.headers.sub_form,
              ],
              // Note 3: Exempted
              [
                { text: l.notes.note3, colSpan: 2 },
                {},
                '3',
                (n['note3']?.val || 0).toLocaleString(),
                '0.00',
                '0.00',
                l.headers.sub_form,
              ],
              // Note 4: Standard Rated
              [
                { text: l.notes.note4, colSpan: 2 },
                {},
                '4',
                (n['note4']?.val || 0).toLocaleString(undefined, { minimumFractionDigits: 2 }),
                (n['note4']?.sd || 0).toLocaleString(),
                (n['note4']?.vat || 0).toLocaleString(undefined, { minimumFractionDigits: 2 }),
                l.headers.sub_form,
              ],
              // Note 5-8: Other Categories
              [
                { text: l.notes.note5, colSpan: 2 },
                {},
                '5',
                (n['note5']?.val || 0).toLocaleString(),
                '0.00',
                '0.00',
                l.headers.sub_form,
              ],
              [
                { text: l.notes.note6, colSpan: 2 },
                {},
                '6',
                (n['note6']?.val || 0).toLocaleString(),
                '0.00',
                '0.00',
                l.headers.sub_form,
              ],
              [
                { text: l.notes.note7, colSpan: 2 },
                {},
                '7',
                (n['note7']?.val || 0).toLocaleString(),
                '0.00',
                '0.00',
                l.headers.sub_form,
              ],
              [
                { text: l.notes.note8, colSpan: 2 },
                {},
                '8',
                (n['note8']?.val || 0).toLocaleString(),
                '0.00',
                '0.00',
                l.headers.sub_form,
              ],
              // Note 9: Total
              [
                { text: l.notes.note9, colSpan: 2, style: 'tBold' },
                {},
                { text: '9', style: 'tBold' },
                {
                  text: (n['note9']?.val || 0).toLocaleString(undefined, {
                    minimumFractionDigits: 2,
                  }),
                  style: 'tBold',
                  fillColor: '#d9d9d9',
                },
                {
                  text: (n['note9']?.sd || 0).toLocaleString(),
                  style: 'tBold',
                  fillColor: '#d9d9d9',
                },
                {
                  text: (n['note9']?.vat || 0).toLocaleString(undefined, {
                    minimumFractionDigits: 2,
                  }),
                  style: 'tBold',
                  fillColor: '#d9d9d9',
                },
                { text: '', border: [false, false, false, false] },
              ],
            ],
          },
        },

        this.createFullWidthHeader(l.sections.s4),
        {
          stack: [
            {
              canvas: [{ type: 'rect', x: 0, y: 0, w: 535, h: 52, color: '#fcd5b4' }],
            },
            {
              text: l.labels.purchase_instruction.join('\n'),
              fontSize: 7,
              margin: [5, -50, 5, 2],
            },
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
                { text: '', border: [false, false, false, false] },
              ],
              // Zero Rated & Exempted (Notes 10-13)
              [
                { text: l.notes.note10, rowSpan: 2 },
                l.labels.local_purchase,
                '10',
                n.note10?.val || '0.00',
                { text: '', fillColor: '#d9d9d9' },
                l.headers.sub_form,
              ],
              [
                {},
                l.labels.import,
                '11',
                n.note11?.val || '0.00',
                { text: '', fillColor: '#d9d9d9' },
                l.headers.sub_form,
              ],
              [
                { text: l.notes.note11, rowSpan: 2 },
                l.labels.local_purchase,
                '12',
                n.note12?.val || '0.00',
                { text: '', fillColor: '#d9d9d9' },
                l.headers.sub_form,
              ],
              [
                {},
                l.labels.import,
                '13',
                n.note13?.val || '0.00',
                { text: '', fillColor: '#d9d9d9' },
                l.headers.sub_form,
              ],

              // Standard Rated - Main Data (Notes 14-15)
              [
                { text: l.notes.note12, rowSpan: 2 },
                l.labels.local_purchase,
                '14',
                n.note14?.val || '0.00',
                n.note14?.vat || '0.00',
                l.headers.sub_form,
              ],
              [
                {},
                l.labels.import,
                '15',
                n.note15?.val || '0.00',
                n.note15?.vat || '0.00',
                l.headers.sub_form,
              ],

              // Other Categories (Notes 16-22)
              [
                { text: l.notes.note13, rowSpan: 2 },
                l.labels.local_purchase,
                '16',
                n.note16?.val || '0.00',
                n.note16?.vat || '0.00',
                l.headers.sub_form,
              ],
              [
                {},
                l.labels.import,
                '17',
                n.note17?.val || '0.00',
                n.note17?.vat || '0.00',
                l.headers.sub_form,
              ],
              [
                { text: l.notes.note14, rowSpan: 1 },
                l.labels.local_purchase,
                '18',
                n.note18?.val || '0.00',
                n.note18?.vat || '0.00',
                l.headers.sub_form,
              ],
              [
                { text: l.notes.note15, rowSpan: 2 },
                l.labels.from_turnover,
                '19',
                n.note19?.val || '0.00',
                n.note19?.vat || '0.00',
                l.headers.sub_form,
              ],
              [
                {},
                l.labels.from_unregistered,
                '20',
                n.note20?.val || '0.00',
                n.note20?.vat || '0.00',
                l.headers.sub_form,
              ],
              [
                { text: l.notes.note16, rowSpan: 2 },
                l.labels.local_purchase,
                '21',
                n.note21?.val || '0.00',
                n.note21?.vat || '0.00',
                l.headers.sub_form,
              ],
              [
                {},
                l.labels.import,
                '22',
                n.note22?.val || n.note22?.val || '0.00',
                n.note22?.vat || '0.00',
                l.headers.sub_form,
              ],

              // Total Row (Note 23)
              [
                { text: l.labels.total_input_credit, colSpan: 1, style: 'tBold', bold: true },
                {},
                { text: '23', style: 'tBold' },
                {
                  text: (n.note23?.val || 3717678.34).toLocaleString(),
                  style: 'tBold',
                  fillColor: '#d9d9d9',
                },
                {
                  text: (n.note23?.vat || 557651.75).toLocaleString(),
                  style: 'tBold',
                  fillColor: '#d9d9d9',
                },
                { text: '', border: [false, false, false, false] },
              ],
            ],
          },
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
                { text: '', style: 'tHead', border: [false, false, false, false] },
              ],
              // Note 24-26
              [
                l.notes.note24,
                { text: '24', alignment: 'center' },
                { text: '0.00', alignment: 'right' },
                l.headers.sub_form,
              ],
              [
                l.notes.note25,
                { text: '25', alignment: 'center' },
                { text: '0.00', alignment: 'right' },
                l.headers.sub_form,
              ],
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
                        body: [[{ text: l.notes.note27_sub, fontSize: 7, bold: true }]],
                      },
                    },
                    // { text: 'VAT on House Rent', margin: [0, 5, 0, 0], bold: true }
                  ],
                },
                { text: '27', alignment: 'center' },
                { text: n.note27?.val || '0.00', alignment: 'right' },
                l.headers.sub_form,
              ],
              // Row 5: Total (Note 28)
              [
                { text: l.labels.total_inc_adj, style: 'tBold', bold: true },
                { text: '28', style: 'tBold', alignment: 'center' },
                { text: n.note28?.val || n.note28 || '0.00', style: 'tBold', alignment: 'right' },
                { text: '', border: [false, false, false, false] },
              ],
            ],
          },
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
                { text: '', style: 'tHead', border: [false, false, false, false] },
              ],
              // Note 29: VDS from supplies delivered
              [
                l.notes.note29,
                { text: '29', alignment: 'center' },
                { text: n.note29?.val || '0.00', alignment: 'right' },
                l.headers.sub_form,
              ],
              // Note 30: Advance Tax
              [
                l.notes.note30,
                { text: '30', alignment: 'center' },
                { text: n.note30?.val || '0.00', alignment: 'right' },
                l.headers.sub_form,
              ],
              // Note 31: Credit Note
              [
                l.notes.note31,
                { text: '31', alignment: 'center' },
                { text: n.note31?.val || '0.00', alignment: 'right' },
                l.headers.sub_form,
              ],
              // Note 32: Other Adjustments with empty box
              [
                {
                  stack: [
                    l.notes.note32,
                    {
                      table: { widths: ['*'], body: [[' ']] },
                      margin: [0, 5, 10, 2],
                    },
                  ],
                },
                { text: '32', alignment: 'center' },
                { text: n.note32?.val || '0.00', alignment: 'right' },
                l.headers.sub_form,
              ],
              // Row 5: Total Decreasing Adjustment (Note 33)
              [
                { text: l.labels.total_dec_adj, style: 'tBold', bold: true },
                { text: '33', style: 'tBold', alignment: 'center' },
                { text: n.note33 || '0.00', style: 'tBold', alignment: 'right' },
                { text: '', border: [false, false, false, false] },
              ],
            ],
          },
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
                { text: l.headers.amount, style: 'tHead', alignment: 'center' },
              ],
              // Notes 34 - 53
              [
                l.notes.note34,
                '34',
                { text: formatAmount(n.note34?.val || n.note34), alignment: 'right' },
              ],
              [
                l.notes.note35,
                '35',
                { text: formatAmount(n.note35?.val || n.note35), alignment: 'right' },
              ],
              [
                l.notes.note36,
                '36',
                { text: formatAmount(n.note36?.val || n.note36), alignment: 'right' },
              ],
              [
                l.notes.note37,
                '37',
                { text: formatAmount(n.note37?.val || n.note37), alignment: 'right' },
              ],
              [
                l.notes.note38,
                '38',
                { text: formatAmount(n.note38?.val || n.note38), alignment: 'right' },
              ],
              [
                l.notes.note39,
                '39',
                { text: formatAmount(n.note39?.val || n.note39), alignment: 'right' },
              ],
              [
                l.notes.note40,
                '40',
                { text: formatAmount(n.note40?.val || n.note40), alignment: 'right' },
              ],
              [
                l.notes.note41,
                '41',
                { text: formatAmount(n.note41?.val || n.note41), alignment: 'right' },
              ],
              [
                l.notes.note42,
                '42',
                { text: formatAmount(n.note42?.val || n.note42), alignment: 'right' },
              ],
              [
                l.notes.note43,
                '43',
                { text: formatAmount(n.note43?.val || n.note43), alignment: 'right' },
              ],
              [
                l.notes.note44,
                '44',
                { text: formatAmount(n.note44?.val || n.note44), alignment: 'right' },
              ],
              [
                l.notes.note45,
                '45',
                { text: formatAmount(n.note45?.val || n.note45), alignment: 'right' },
              ],
              [
                l.notes.note46,
                '46',
                { text: formatAmount(n.note46?.val || n.note46), alignment: 'right' },
              ],
              [
                l.notes.note47,
                '47',
                { text: formatAmount(n.note47?.val || n.note47), alignment: 'right' },
              ],
              [
                l.notes.note48,
                '48',
                { text: formatAmount(n.note48?.val || n.note48), alignment: 'right' },
              ],
              [
                l.notes.note49,
                '49',
                { text: formatAmount(n.note49?.val || n.note49), alignment: 'right' },
              ],
              [
                l.notes.note50,
                '50',
                { text: formatAmount(n.note50?.val || n.note50), alignment: 'right' },
              ],
              [
                l.notes.note51,
                '51',
                { text: formatAmount(n.note51?.val || n.note51), alignment: 'right' },
              ],
              [
                l.notes.note52,
                '52',
                { text: formatAmount(n.note52?.val || n.note52), alignment: 'right' },
              ],
              [
                l.notes.note53,
                '53',
                { text: formatAmount(n.note53?.val || n.note53), alignment: 'right' },
              ],
            ],
          },
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
                { text: l.headers.amount, style: 'tHead', alignment: 'center' },
              ],
              // Notes 54 - 57
              [
                l.notes.note54,
                { text: '54', alignment: 'center' },
                { text: (n.note54 || '0.00').toLocaleString(), alignment: 'right' },
              ],
              [
                l.notes.note55,
                { text: '55', alignment: 'center' },
                { text: n.note55?.val || n.note55 || '0.00', alignment: 'right' },
              ],
              [
                l.notes.note56,
                { text: '56', alignment: 'center' },
                { text: n.note56?.val || n.note56 || '0.00', alignment: 'right' },
              ],
              [
                l.notes.note57,
                { text: '57', alignment: 'center' },
                { text: n.note57?.val || n.note57 || '0.00', alignment: 'right' },
              ],
            ],
          },
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
                { text: '', style: 'tHead' },
              ],
              // Row 58: VAT Deposit
              [
                l.notes.note58,
                '58',
                n.note58?.code || '1/1133/0030/0311',
                n.note58?.val || '0.00',
                l.headers.sub_form,
              ],
              // Row 59: SD Deposit
              [
                l.notes.note59,
                '59',
                n.note59?.code || '1/1133/0018/0711-0721',
                n.note59?.val || '0.00',
                l.headers.sub_form,
              ],
              // Row 60: Excise Duty
              [
                l.notes.note60,
                '60',
                n.note60?.code || '1/1133/Acv‡ikbvj †KvW/0311',
                n.note60?.val || '0.00',
                l.headers.sub_form,
              ],
              // Row 61: Development Surcharge
              [
                l.notes.note61,
                '61',
                n.note61?.code || '1/1133/Acv‡ikbvj',
                n.note61?.val || '0.00',
                l.headers.sub_form,
              ],
              // Row 62: ICT Development Surcharge
              [
                l.notes.note62,
                '62',
                n.note62?.code || '1/1103/Acv‡ikbvj †KvW/1901',
                n.note62?.val || '0.00',
                l.headers.sub_form,
              ],
              // Row 63: Health Care Surcharge
              [
                l.notes.note63,
                '63',
                n.note63?.code || '1/1133/Acv‡ikbvj †KvW/0601',
                n.note63?.val || '0.00',
                l.headers.sub_form,
              ],
              // Row 64: Environmental Protection Surcharge
              [
                l.notes.note64,
                '64',
                n.note64?.code || '1/1103/Acv‡ikbvj †KvW/2225',
                n.note64?.val || '0.00',
                l.headers.sub_form,
              ],
            ],
          },
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
                { text: l.headers.amount, style: 'tHead', alignment: 'center' },
              ],
              // Row 65: Closing Balance (VAT)
              [
                l.notes.note65,
                { text: '65', alignment: 'center' },
                { text: n.note65?.val || n.note65 || '0.00', alignment: 'right', style: 'tBold' },
              ],
              // Row 66: Closing Balance (SD)
              [
                l.notes.note66,
                { text: '66', alignment: 'center' },
                { text: n.note66?.val || n.note66 || '0.00', alignment: 'right' },
              ],
            ],
          },
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
                    { width: 'auto', text: l.labels.no, fontSize: 7, margin: [0, 2, 0, 0] },
                  ],
                  style: 'tHead',
                },
              ],
              // Note 67
              [
                {},
                l.labels.req_refund_vat,
                { text: '67', alignment: 'center' },
                { text: n.note67?.val || n.note67 || '0.00', alignment: 'right' },
              ],
              // Note 68
              [
                {},
                l.labels.req_refund_sd,
                { text: '68', alignment: 'center' },
                { text: n.note68?.val || n.note68 || '0.00', alignment: 'right' },
              ],
            ],
          },
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
                  margin: [5, 5, 5, 5],
                },
              ],
            ],
          },
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
              [l.labels.signature, ':', ''],
            ],
          },
        },
      ],
      styles: {
        header: { font: 'PlaywriteCU', fontSize: 8, bold: true, alignment: 'center' },
        subHeader: {
          font: 'PlaywriteCU',
          fontSize: 7,
          bold: true,
          alignment: 'center',
          color: '#003366',
        },
        secHeaderCell: {
          font: 'PlaywriteCU',
          fillColor: '#003366',
          color: 'white',
          bold: true,
          alignment: 'center',
          fontSize: 7,
          padding: [0, 2, 0, 2],
        },
        tHead: { font: 'PlaywriteCU', fillColor: '#f2f2f2', bold: true, fontSize: 7 },
        tBold: { font: 'PlaywriteCU', bold: true, fontSize: 7 },
        dataTable: { font: 'PlaywriteCU', fontSize: 7, margin: [0, 0, 0, 5] },
        borderedTable: { font: 'PlaywriteCU', margin: [0, 0, 0, 2] },
      },
    };
    pdfMake.createPdf(docDef).download(`Mushak_9.1_${lang}.pdf`);
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

    const brandRow2 = sheet.addRow(['NATIONAL BOARD OF REVENUE', '', '']);
    sheet.mergeCells(`B${brandRow2.number}:E${brandRow2.number}`);
    brandRow2.getCell(1).font = { size: 11, bold: true };
    brandRow2.getCell(1).alignment = { horizontal: 'center' };

    const formTitleRow = sheet.addRow(['VALUE ADDED TAX RETURN FORM', '', '|| Mushak-9.1 ||']);
    sheet.mergeCells(`B${formTitleRow.number}:E${formTitleRow.number}`);
    formTitleRow.getCell(1).font = { size: 10, bold: true };
    formTitleRow.getCell(1).alignment = { horizontal: 'center' };
    formTitleRow.getCell(6).font = { size: 10, bold: true }; // Mushak-9.1 ID on right

    const ruleRow = sheet.addRow(['[Rule 47(1)]', '', '']);
    sheet.mergeCells(`B${ruleRow.number}:E${ruleRow.number}`);
    ruleRow.getCell(1).font = { size: 8 };
    ruleRow.getCell(1).alignment = { horizontal: 'center' };

    // --- 1. COLUMN SETUP ---
    sheet.columns = [
      { width: 35 }, // A: Label
      { width: 3 }, // B: Separator (:)
      { width: 35 }, // C: Data
      { width: 10 }, // D: Spacing
      { width: 10 }, // E: Spacing
      { width: 12 }, // F: সাবফর্ম
    ];

    // --- 2. STYLING HELPERS ---
    const addHeader = (text: string) => {
      sheet.addRow([]); // Spacer
      const row = sheet.addRow([text]);
      sheet.mergeCells(`A${row.number}:E${row.number}`);
      row.eachCell((c) => {
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
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' },
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
      ['5. Economic Activity', ':', data.taxpayer.activity],
    ];
    s1Data.forEach((item) => {
      const r = sheet.addRow([item[0], item[1], item[2]]);
      sheet.mergeCells(`C${r.number}:E${r.number}`); // Merged for full-width data row
    });
    applyBorder(s1Start, sheet.rowCount);

    addHeader('SECTION - 2: RETURN SUBMISSION DATA');
    const s2Start = sheet.rowCount + 1;
    const s2Data = [
      ['1. Tax Period', ':', data.returnSubmission.period],
      ['2. Type of Return', ':', 'A) Main/Original Return (Section 64)   [ X ]'],
      ['', '', 'B) Late Return (section 65)   [   ]'],
      ['', '', 'C) Amend Return (section 66)   [   ]'],
      ['3. Any activities in this Tax Period?', ':', '[ X ] Yes     [   ] No'],
      ['4. Date of Submission', ':', data.returnSubmission.date],
    ];
    s2Data.forEach((item) => {
      const r = sheet.addRow([item[0], item[1], item[2]]);
      sheet.mergeCells(`C${r.number}:E${r.number}`); // Merged for full-width data row
    });
    applyBorder(s2Start, sheet.rowCount);

    // --- SECTION 3 & 4 ---
    addHeader('SECTION - 3: SUPPLY - OUTPUT TAX');
    const s3Start = sheet.rowCount + 1;
    // Header Row
    const head = sheet.addRow([
      'Nature of Supply',
      '',
      'Note',
      'Value (a)',
      'SD (b)',
      'VAT (c)',
      '',
    ]);
    sheet.mergeCells(`A${head.number}:B${head.number}`);
    head.font = { bold: true };

    // Data Rows (Standard Rated & Total)
    sheet.addRow(['Zero Rated Goods/Service', 'Direct Export', '1', '', '', '', 'সাবফর্ম']);
    sheet.addRow(['', 'Deemed Export', '2', '', '', '', 'সাবফর্ম']);
    sheet.addRow(['Exempted Goods/Service', '', '3', '', '', '', 'সাবফর্ম']);

    // Note 4
    const n4 = sheet.addRow([
      'Standard Rated Goods/Service',
      '',
      '4',
      159270.3,
      0,
      23890.55,
      'সাবফর্ম',
    ]);
    sheet.mergeCells(`A${n4.number}:B${n4.number}`);

    sheet.addRow(['Goods Based on MRP', '', '5', '', '', '', 'সাবফর্ম']);
    sheet.addRow(['Goods/Service Based on Specific VAT', '', '6', '', '', '', 'সাবফর্ম']);
    sheet.addRow(['Goods/Service Other than Standard Rate', '', '7', '', '', '', 'সাবফর্ম']);
    sheet.addRow(['Retail/Whole Sale/Trade Based Supply', '', '8', 0, 0, 0, 'সাবফর্ম']);

    // Note 9 Total
    const n9 = sheet.addRow([
      'Total Sales Value & Total Payable Taxes',
      '',
      '9',
      159270.3,
      0,
      23890.55,
      '',
    ]);
    sheet.mergeCells(`A${n9.number}:B${n9.number}`);
    n9.eachCell((c) => {
      c.font = { bold: true };
      if (c.address.includes('D') || c.address.includes('E') || c.address.includes('F')) {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9D9D9' } };
      }
    });
    applyBorder(s3Start, sheet.rowCount);

    addHeader('SECTION - 4: PURCHASE - INPUT TAX');
    const s4Start = sheet.rowCount + 1;
    sheet.addRow(['Nature', 'Note', 'Value', 'VAT', 'Remarks']).font = { bold: true };
    sheet.addRow(['Local Purchase', '14', data.notes.note14.val, data.notes.note14.vat, 'সাবফর্ম']);
    applyBorder(s4Start, sheet.rowCount);

    // --- INDIVIDUAL SECTIONS 5, 6, 7, 8 (FIXED) ---
    addHeader('SECTION - 5: INCREASING ADJUSTMENTS');
    const s5Start = sheet.rowCount + 1;
    sheet.addRow(['Total Increasing Adjustment', '28', '', data.notes.note28]);
    applyBorder(s5Start, sheet.rowCount);

    addHeader('SECTION - 6: DECREASING ADJUSTMENTS');
    const s6Start = sheet.rowCount + 1;
    sheet.addRow(['Total Decreasing Adjustment', '33', '', data.notes.note33]);
    applyBorder(s6Start, sheet.rowCount);

    addHeader('SECTION - 7: NET TAX CALCULATION');
    const s7Start = sheet.rowCount + 1;
    sheet.addRow(['Net Payable VAT (34)', '34', '', data.notes.note34]);
    sheet.addRow(['Payable for Treasury Deposit', '50', '', data.notes.note50]);
    applyBorder(s7Start, sheet.rowCount);

    addHeader('SECTION - 8: OLD ACCOUNT BALANCE');
    const s8Start = sheet.rowCount + 1;
    sheet.addRow(['Balance from Mushak-18.6', '54', '', data.notes.note54]);
    applyBorder(s8Start, sheet.rowCount);

    // --- INDIVIDUAL SECTIONS 9, 10, 11, 12 ---
    addHeader('SECTION - 9: ACCOUNT CODE WISE PAYMENT SCHEDULE');
    const s9Start = sheet.rowCount + 1;
    sheet.addRow(['VAT Deposit', '58', '1/1133/0030/0311', '0.00']);
    applyBorder(s9Start, sheet.rowCount);

    addHeader('SECTION - 10: CLOSING BALANCE');
    const s10Start = sheet.rowCount + 1;
    sheet.addRow(['Closing Balance (VAT)', '65', '', data.notes.note65]);
    applyBorder(s10Start, sheet.rowCount);

    addHeader('SECTION - 11: REFUND');
    const s11Start = sheet.rowCount + 1;
    sheet.addRow(['Requested Refund (VAT)', '67', '', data.notes.note67]);
    applyBorder(s11Start, sheet.rowCount);

    addHeader('SECTION - 12: DECLARATION');
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
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
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
        fontSize: 7,
      },
      content: [
        {
          stack: [
            { text: l.titles.gov, style: 'header' },
            { text: l.titles.nbr, style: 'header' },
            { text: `\n${l.titles.form}`, style: 'subHeader' },
            { text: `${l.titles.rule}\n`, style: 'subHeader' },
            { text: '\n', style: 'subHeader' },
          ],
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
              [l.labels.activity, ':', t.activity],
            ],
          },
        },

        this.createFullWidthHeader(l.sections.s2),
        {
          style: 'dataTable',
          table: {
            widths: ['35%', '2%', '63%'],
            body: [
              [l.labels.tax_period, ':', { text: s.period || 'Oct / 2022', alignment: 'center' }],
              [
                l.labels.return_type,
                ':',
                {
                  stack: [
                    {
                      columns: [
                        { width: '70%', text: l.return_options ? l.return_options[0] : '' },
                        {
                          table: { widths: ['30%'], body: [[' ']] },
                          margin: [0, 0, 10, 2],
                          alignment: 'right',
                        },
                      ],
                    },
                    {
                      columns: [
                        { width: '70%', text: l.return_options ? l.return_options[1] : '' },
                        {
                          table: { widths: ['30%'], body: [[' ']] },
                          margin: [0, 0, 10, 2],
                          alignment: 'right',
                        },
                      ],
                    },
                    {
                      columns: [
                        { width: '70%', text: l.return_options ? l.return_options[2] : '' },
                        {
                          table: { widths: ['30%'], body: [[' ']] },
                          margin: [0, 0, 10, 2],
                          alignment: 'right',
                        },
                      ],
                    },
                    {
                      columns: [
                        { width: '70%', text: l.return_options ? l.return_options[3] : '' },
                        {
                          table: { widths: ['30%'], body: [[' ']] },
                          margin: [0, 0, 10, 2],
                          alignment: 'right',
                        },
                      ],
                    },
                  ],
                  margin: [0, 2, 0, 1],
                },
              ],
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
                            {
                              width: 'auto',
                              table: { widths: [20], body: [[' ']] },
                              margin: [0, 0, 5, 0],
                            },
                            {
                              width: 'auto',
                              text: l.labels.yes,
                              fontSize: 7,
                              margin: [0, 2, 25, 0],
                            },

                            // No Option
                            {
                              width: 'auto',
                              table: { widths: [20], body: [[' ']] },
                              margin: [0, 0, 5, 0],
                            },
                            { width: 'auto', text: l.labels.no, fontSize: 7, margin: [0, 2, 0, 0] },
                          ],
                        },
                        { width: '*', text: '' },
                      ],
                    },
                    {
                      text: l.labels.activity_note,
                      fontSize: 7,
                      alignment: 'center',
                      margin: [0, 5, 0, 0],
                      color: '#333333',
                    },
                  ],
                  margin: [0, 2, 0, 1],
                },
              ],
              [l.labels.sub_date, ':', { text: s.date || 'Oct / 2022', alignment: 'center' }],
            ],
          },
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
                { text: '', border: [false, false, false, false] },
              ],
              // Note 1 & 2: Zero Rated
              [
                { text: l.notes.note1.split('-')[0], rowSpan: 2 },
                l.notes.note1.split('-')[1] || l.notes.note1,
                '১',
                (n['note1']?.val || 0).toLocaleString(),
                (n['note1']?.sd || 0).toLocaleString(),
                (n['note1']?.vat || 0).toLocaleString(),
                l.headers.sub_form,
              ],
              [
                {},
                l.notes.note2.split('-')[1] || l.notes.note2,
                '২',
                (n['note2']?.val || 0).toLocaleString(),
                (n['note2']?.sd || 0).toLocaleString(),
                (n['note2']?.vat || 0).toLocaleString(),
                l.headers.sub_form,
              ],
              // Note 3: Exempted
              [
                { text: l.notes.note3, colSpan: 2 },
                {},
                '৩',
                (n['note3']?.val || 0).toLocaleString(),
                '0.00',
                '0.00',
                l.headers.sub_form,
              ],
              // Note 4: Standard Rated
              [
                { text: l.notes.note4, colSpan: 2 },
                {},
                '৪',
                (n['note4']?.val || 0).toLocaleString(undefined, { minimumFractionDigits: 2 }),
                (n['note4']?.sd || 0).toLocaleString(),
                (n['note4']?.vat || 0).toLocaleString(undefined, { minimumFractionDigits: 2 }),
                l.headers.sub_form,
              ],
              // Note 5-8: Other Categories
              [
                { text: l.notes.note5, colSpan: 2 },
                {},
                '৫',
                (n['note5']?.val || 0).toLocaleString(),
                '0.00',
                '0.00',
                l.headers.sub_form,
              ],
              [
                { text: l.notes.note6, colSpan: 2 },
                {},
                '৬',
                (n['note6']?.val || 0).toLocaleString(),
                '0.00',
                '0.00',
                l.headers.sub_form,
              ],
              [
                { text: l.notes.note7, colSpan: 2 },
                {},
                '৭',
                (n['note7']?.val || 0).toLocaleString(),
                '0.00',
                '0.00',
                l.headers.sub_form,
              ],
              [
                { text: l.notes.note8, colSpan: 2 },
                {},
                '৮',
                (n['note8']?.val || 0).toLocaleString(),
                '0.00',
                '0.00',
                l.headers.sub_form,
              ],
              // Note 9: Total
              [
                { text: l.notes.note9, colSpan: 2, style: 'tBold' },
                {},
                { text: '৯', style: 'tBold' },
                {
                  text: (n['note9']?.val || 0).toLocaleString(undefined, {
                    minimumFractionDigits: 2,
                  }),
                  style: 'tBold',
                  fillColor: '#d9d9d9',
                },
                {
                  text: (n['note9']?.sd || 0).toLocaleString(),
                  style: 'tBold',
                  fillColor: '#d9d9d9',
                },
                {
                  text: (n['note9']?.vat || 0).toLocaleString(undefined, {
                    minimumFractionDigits: 2,
                  }),
                  style: 'tBold',
                  fillColor: '#d9d9d9',
                },
                { text: '', border: [false, false, false, false] },
              ],
            ],
          },
        },

        this.createFullWidthHeader(l.sections.s4),
        {
          stack: [
            {
              canvas: [{ type: 'rect', x: 0, y: 0, w: 535, h: 42, color: '#fcd5b4' }],
            },
            {
              text: l.labels.purchase_instruction.join('\n'),
              fontSize: 7,
              margin: [5, -38, 5, 2],
            },
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
                {
                  text: l.headers.nature_purchase,
                  style: 'tHead',
                  colSpan: 2,
                  alignment: 'center',
                },
                {},
                { text: l.headers.note, style: 'tHead', alignment: 'center' },
                { text: l.headers.value, style: 'tHead', alignment: 'center' },
                { text: l.headers.vat, style: 'tHead', alignment: 'center' },
                { text: '', border: [false, false, false, false] },
              ],
              // Zero Rated & Exempted (Notes 10-13)
              [
                { text: l.notes.note10, rowSpan: 2 },
                l.labels.local_purchase,
                '10',
                n.note10?.val || '0.00',
                { text: '', fillColor: '#d9d9d9' },
                l.headers.sub_form,
              ],
              [
                {},
                l.labels.import,
                '11',
                n.note11?.val || '0.00',
                { text: '', fillColor: '#d9d9d9' },
                l.headers.sub_form,
              ],
              [
                { text: l.notes.note11, rowSpan: 2 },
                l.labels.local_purchase,
                '12',
                n.note12?.val || '0.00',
                { text: '', fillColor: '#d9d9d9' },
                l.headers.sub_form,
              ],
              [
                {},
                l.labels.import,
                '13',
                n.note13?.val || '0.00',
                { text: '', fillColor: '#d9d9d9' },
                l.headers.sub_form,
              ],

              // Standard Rated - Main Data (Notes 14-15)
              [
                { text: l.notes.note12, rowSpan: 2 },
                l.labels.local_purchase,
                '14',
                n.note14?.val || '0.00',
                n.note14?.vat || '0.00',
                l.headers.sub_form,
              ],
              [
                {},
                l.labels.import,
                '15',
                n.note15?.val || '0.00',
                n.note15?.vat || '0.00',
                l.headers.sub_form,
              ],

              // Other Categories (Notes 16-22)
              [
                { text: l.notes.note13, rowSpan: 2 },
                l.labels.local_purchase,
                '16',
                n.note16?.val || '0.00',
                n.note16?.vat || '0.00',
                l.headers.sub_form,
              ],
              [
                {},
                l.labels.import,
                '17',
                n.note17?.val || '0.00',
                n.note17?.vat || '0.00',
                l.headers.sub_form,
              ],
              [
                { text: l.notes.note14, rowSpan: 1 },
                l.labels.local_purchase,
                '18',
                n.note18?.val || '0.00',
                n.note18?.vat || '0.00',
                l.headers.sub_form,
              ],
              [
                { text: l.notes.note15, rowSpan: 2 },
                l.labels.from_turnover,
                '19',
                n.note19?.val || '0.00',
                n.note19?.vat || '0.00',
                l.headers.sub_form,
              ],
              [
                {},
                l.labels.from_unregistered,
                '20',
                n.note20?.val || '0.00',
                n.note20?.vat || '0.00',
                l.headers.sub_form,
              ],
              [
                { text: l.notes.note16, rowSpan: 2 },
                l.labels.local_purchase,
                '21',
                n.note21?.val || '0.00',
                n.note21?.vat || '0.00',
                l.headers.sub_form,
              ],
              [
                {},
                l.labels.import,
                '22',
                n.note22?.val || n.note22?.val || '0.00',
                n.note22?.vat || '0.00',
                l.headers.sub_form,
              ],

              // Total Row (Note 23)
              [
                { text: l.labels.total_input_credit, colSpan: 1, style: 'tBold', bold: true },
                {},
                { text: '23', style: 'tBold' },
                {
                  text: (n.note23?.val || 3717678.34).toLocaleString(),
                  style: 'tBold',
                  fillColor: '#d9d9d9',
                },
                {
                  text: (n.note23?.vat || 557651.75).toLocaleString(),
                  style: 'tBold',
                  fillColor: '#d9d9d9',
                },
                { text: '', border: [false, false, false, false] },
              ],
            ],
          },
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
                { text: '', style: 'tHead', border: [false, false, false, false] },
              ],
              // Note 24-26
              [
                l.notes.note24,
                { text: '24', alignment: 'center' },
                { text: '0.00', alignment: 'right' },
                l.headers.sub_form,
              ],
              [
                l.notes.note25,
                { text: '25', alignment: 'center' },
                { text: '0.00', alignment: 'right' },
                l.headers.sub_form,
              ],
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
                        body: [[{ text: l.notes.note27_sub, fontSize: 7, bold: true }]],
                      },
                    },
                    // { text: 'VAT on House Rent', margin: [0, 5, 0, 0], bold: true }
                  ],
                },
                { text: '27', alignment: 'center' },
                { text: n.note27?.val || '0.00', alignment: 'right' },
                l.headers.sub_form,
              ],
              // Row 5: Total (Note 28)
              [
                { text: l.labels.total_inc_adj, style: 'tBold', bold: true },
                { text: '28', style: 'tBold', alignment: 'center' },
                { text: n.note28?.val || n.note28 || '0.00', style: 'tBold', alignment: 'right' },
                { text: '', border: [false, false, false, false] },
              ],
            ],
          },
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
                { text: '', style: 'tHead', border: [false, false, false, false] },
              ],
              // Note 29: VDS from supplies delivered
              [
                l.notes.note29,
                { text: '29', alignment: 'center' },
                { text: n.note29?.val || '0.00', alignment: 'right' },
                l.headers.sub_form,
              ],
              // Note 30: Advance Tax
              [
                l.notes.note30,
                { text: '30', alignment: 'center' },
                { text: n.note30?.val || '0.00', alignment: 'right' },
                l.headers.sub_form,
              ],
              // Note 31: Credit Note
              [
                l.notes.note31,
                { text: '31', alignment: 'center' },
                { text: n.note31?.val || '0.00', alignment: 'right' },
                l.headers.sub_form,
              ],
              // Note 32: Other Adjustments with empty box
              [
                {
                  stack: [
                    l.notes.note32,
                    {
                      table: { widths: ['*'], body: [[' ']] },
                      margin: [0, 5, 10, 2],
                    },
                  ],
                },
                { text: '32', alignment: 'center' },
                { text: n.note32?.val || '0.00', alignment: 'right' },
                l.headers.sub_form,
              ],
              // Row 5: Total Decreasing Adjustment (Note 33)
              [
                { text: l.labels.total_dec_adj, style: 'tBold', bold: true },
                { text: '33', style: 'tBold', alignment: 'center' },
                { text: n.note33 || '0.00', style: 'tBold', alignment: 'right' },
                { text: '', border: [false, false, false, false] },
              ],
            ],
          },
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
                { text: l.headers.amount, style: 'tHead', alignment: 'center' },
              ],
              // Notes 34 - 53
              [
                l.notes.note34,
                '34',
                { text: formatAmount(n.note34?.val || n.note34), alignment: 'right' },
              ],
              [
                l.notes.note35,
                '35',
                { text: formatAmount(n.note35?.val || n.note35), alignment: 'right' },
              ],
              [
                l.notes.note36,
                '36',
                { text: formatAmount(n.note36?.val || n.note36), alignment: 'right' },
              ],
              [
                l.notes.note37,
                '37',
                { text: formatAmount(n.note37?.val || n.note37), alignment: 'right' },
              ],
              [
                l.notes.note38,
                '38',
                { text: formatAmount(n.note38?.val || n.note38), alignment: 'right' },
              ],
              [
                l.notes.note39,
                '39',
                { text: formatAmount(n.note39?.val || n.note39), alignment: 'right' },
              ],
              [
                l.notes.note40,
                '40',
                { text: formatAmount(n.note40?.val || n.note40), alignment: 'right' },
              ],
              [
                l.notes.note41,
                '41',
                { text: formatAmount(n.note41?.val || n.note41), alignment: 'right' },
              ],
              [
                l.notes.note42,
                '42',
                { text: formatAmount(n.note42?.val || n.note42), alignment: 'right' },
              ],
              [
                l.notes.note43,
                '43',
                { text: formatAmount(n.note43?.val || n.note43), alignment: 'right' },
              ],
              [
                l.notes.note44,
                '44',
                { text: formatAmount(n.note44?.val || n.note44), alignment: 'right' },
              ],
              [
                l.notes.note45,
                '45',
                { text: formatAmount(n.note45?.val || n.note45), alignment: 'right' },
              ],
              [
                l.notes.note46,
                '46',
                { text: formatAmount(n.note46?.val || n.note46), alignment: 'right' },
              ],
              [
                l.notes.note47,
                '47',
                { text: formatAmount(n.note47?.val || n.note47), alignment: 'right' },
              ],
              [
                l.notes.note48,
                '48',
                { text: formatAmount(n.note48?.val || n.note48), alignment: 'right' },
              ],
              [
                l.notes.note49,
                '49',
                { text: formatAmount(n.note49?.val || n.note49), alignment: 'right' },
              ],
              [
                l.notes.note50,
                '50',
                { text: formatAmount(n.note50?.val || n.note50), alignment: 'right' },
              ],
              [
                l.notes.note51,
                '51',
                { text: formatAmount(n.note51?.val || n.note51), alignment: 'right' },
              ],
              [
                l.notes.note52,
                '52',
                { text: formatAmount(n.note52?.val || n.note52), alignment: 'right' },
              ],
              [
                l.notes.note53,
                '53',
                { text: formatAmount(n.note53?.val || n.note53), alignment: 'right' },
              ],
            ],
          },
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
                { text: l.headers.amount, style: 'tHead', alignment: 'center' },
              ],
              // Notes 54 - 57
              [
                l.notes.note54,
                { text: '54', alignment: 'center' },
                { text: (n.note54 || '0.00').toLocaleString(), alignment: 'right' },
              ],
              [
                l.notes.note55,
                { text: '55', alignment: 'center' },
                { text: n.note55?.val || n.note55 || '0.00', alignment: 'right' },
              ],
              [
                l.notes.note56,
                { text: '56', alignment: 'center' },
                { text: n.note56?.val || n.note56 || '0.00', alignment: 'right' },
              ],
              [
                l.notes.note57,
                { text: '57', alignment: 'center' },
                { text: n.note57?.val || n.note57 || '0.00', alignment: 'right' },
              ],
            ],
          },
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
                { text: '', style: 'tHead' },
              ],
              // Row 58: VAT Deposit
              [
                l.notes.note58,
                '58',
                n.note58?.code || '1/1133/0030/0311',
                n.note58?.val || '0.00',
                l.headers.sub_form,
              ],
              // Row 59: SD Deposit
              [
                l.notes.note59,
                '59',
                n.note59?.code || '1/1133/0018/0711-0721',
                n.note59?.val || '0.00',
                l.headers.sub_form,
              ],
              // Row 60: Excise Duty
              [
                l.notes.note60,
                '60',
                n.note60?.code || '1/1133/Acv‡ikbvj †KvW/0311',
                n.note60?.val || '0.00',
                l.headers.sub_form,
              ],
              // Row 61: Development Surcharge
              [
                l.notes.note61,
                '61',
                n.note61?.code || '1/1133/Acv‡ikbvj',
                n.note61?.val || '0.00',
                l.headers.sub_form,
              ],
              // Row 62: ICT Development Surcharge
              [
                l.notes.note62,
                '62',
                n.note62?.code || '1/1103/Acv‡ikbvj †KvW/1901',
                n.note62?.val || '0.00',
                l.headers.sub_form,
              ],
              // Row 63: Health Care Surcharge
              [
                l.notes.note63,
                '63',
                n.note63?.code || '1/1133/Acv‡ikbvj †KvW/0601',
                n.note63?.val || '0.00',
                l.headers.sub_form,
              ],
              // Row 64: Environmental Protection Surcharge
              [
                l.notes.note64,
                '64',
                n.note64?.code || '1/1103/Acv‡ikbvj †KvW/2225',
                n.note64?.val || '0.00',
                l.headers.sub_form,
              ],
            ],
          },
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
                { text: l.headers.amount, style: 'tHead', alignment: 'center' },
              ],
              // Row 65: Closing Balance (VAT)
              [
                l.notes.note65,
                { text: '65', alignment: 'center' },
                { text: n.note65?.val || n.note65 || '0.00', alignment: 'right', style: 'tBold' },
              ],
              // Row 66: Closing Balance (SD)
              [
                l.notes.note66,
                { text: '66', alignment: 'center' },
                { text: n.note66?.val || n.note66 || '0.00', alignment: 'right' },
              ],
            ],
          },
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
                    { width: 'auto', text: l.labels.no, fontSize: 7, margin: [0, 2, 0, 0] },
                  ],
                  style: 'tHead',
                },
              ],
              // Note 67
              [
                {},
                l.labels.req_refund_vat,
                { text: '67', alignment: 'center' },
                { text: n.note67?.val || n.note67 || '0.00', alignment: 'right' },
              ],
              // Note 68
              [
                {},
                l.labels.req_refund_sd,
                { text: '68', alignment: 'center' },
                { text: n.note68?.val || n.note68 || '0.00', alignment: 'right' },
              ],
            ],
          },
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
                  margin: [5, 5, 5, 5],
                },
              ],
            ],
          },
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
              [l.labels.signature, ':', ''],
            ],
          },
        },
      ],
      styles: {
        header: { font: 'PlaywriteCU', fontSize: 8, bold: true, alignment: 'center' },
        subHeader: {
          font: 'PlaywriteCU',
          fontSize: 7,
          bold: true,
          alignment: 'center',
          color: '#003366',
        },
        secHeaderCell: {
          font: 'PlaywriteCU',
          fillColor: '#003366',
          color: 'white',
          bold: true,
          alignment: 'center',
          fontSize: 7,
          padding: [0, 2, 0, 2],
        },
        tHead: { font: 'PlaywriteCU', fillColor: '#f2f2f2', bold: true, fontSize: 7 },
        tBold: { font: 'PlaywriteCU', bold: true, fontSize: 7 },
        dataTable: { font: 'PlaywriteCU', fontSize: 7, margin: [0, 0, 0, 5] },
        borderedTable: { font: 'PlaywriteCU', margin: [0, 0, 0, 2] },
      },
    };
    pdfMake.createPdf(docDef).download(`Mushak_9.1_${lang}.pdf`);
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
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
    };

    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');

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
                { text: l.titles.rule, style: 'subHeader' },
              ],
              width: 400,
            },
            { text: l.titles.m_name, alignment: 'right', bold: true, fontSize: 12, width: '*' },
          ],
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
              [l.info.first_supply, ':', safe(info.firstSupplyDate)],
            ],
          },
          layout: 'noBorders',
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
                {
                  text: 'Description of Raw Materials, Quantity & Purchase Price',
                  colSpan: 5,
                  alignment: 'center',
                  bold: true,
                },
                {},
                {},
                {},
                {},
                { text: 'Value Addition Details', colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: l.headers.remarks, rowSpan: 2, alignment: 'center', bold: true },
              ],
              // Row 2: Sub-headers
              [
                {},
                {},
                {},
                {},
                { text: l.headers.raw_material, bold: true },
                { text: l.headers.buy_price, bold: true },
                { text: l.headers.qty_w, bold: true },
                { text: l.headers.qty_wo, bold: true },
                { text: l.headers.wastage_p, bold: true },
                { text: l.headers.va_sector, bold: true },
                { text: l.headers.va_value, bold: true },
                {},
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
                safe(item.remarks),
              ]),
            ],
          },
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
                { text: f.seal, margin: [0, 5, 0, 0] },
              ],
              width: 280,
            },
          ],
        },
        {
          stack: [
            {
              text: f.special_note_title,
              bold: true,
              decoration: 'underline',
              margin: [0, 10, 0, 5],
            },
            {
              ol: f.notes || [],
              fontSize: 7,
              lineHeight: 1.3,
            },
          ],
        },
      ],
      styles: {
        header: { fontSize: 11, bold: true, alignment: 'center' },
        subHeader: { fontSize: 9, alignment: 'center' },
      },
    };

    pdfMake.createPdf(docDef).download(`mushak_4_3_${lang}.pdf`);
  }

  exportInputOutputCoefficientBangla(data: any, lang: string) {
    const l = (data.labels.mushak_4_3 || {}) as any;
    const info = data.mushak_4_3_data[lang]?.companyInfo || {};
    const items = (data.mushak_4_3_data[lang]?.items || []) as any[];
    const f = l.footer || {};

    (pdfMake as any).fonts = {
      kalpurush: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
      },
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
                { text: l.titles.rule, style: 'subHeader' },
              ],
              width: 300,
            },
            { text: l.titles.m_name, alignment: 'right', bold: true, fontSize: 12, width: '*' },
          ],
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
              [l.info.first_supply, ':', safeText(info.firstSupplyDate, true)],
            ],
          },
          layout: 'noBorders',
        },

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
                {},
                {},
                {},
                {},
                { text: l.headers.price_correction, colSpan: 2, alignment: 'center' },
                {},
                { text: l.headers.remarks, rowSpan: 2, alignment: 'center' },
              ],
              [
                {},
                {},
                {},
                {},
                l.headers.raw_material,
                l.headers.buy_price,
                l.headers.qty_w,
                l.headers.qty_wo,
                l.headers.wastage_p,
                l.headers.va_sector,
                l.headers.va_value,
                {},
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
                safeText(item.remarks),
              ]),
            ],
          },
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
                { text: f.seal, margin: [0, 5, 0, 0] },
              ],
              width: 250,
              alignment: 'left',
            },
          ],
        },

        {
          stack: [
            {
              text: f.special_note_title,
              bold: true,
              decoration: 'underline',
              margin: [0, 10, 0, 5],
            },
            {
              text: (f.notes || []).join('\n'),
              fontSize: 7,
              lineHeight: 1.4,
            },
          ],
        },
      ],
      styles: {
        header: { fontSize: 11, bold: true, alignment: 'center' },
        subHeader: { fontSize: 9, alignment: 'center' },
      },
    };

    pdfMake.createPdf(docDef).download(`mushak_4_3_${lang}.pdf`);
  }

  // Mushak 2.3 Export Function
  exportMushak_2_3(data: any, lang: string) {
    const l = (data.labels?.mushak_2_3 || {}) as any;
    const targetData = data.mushak_2_3_data?.[lang] || {};
    const d = targetData;

    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : 'N/A');

    (pdfMake as any).fonts = {
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
    };

    // Generate QR code as base64 using qrcode library
    // const QRCode = require('qrcode');
    QRCode.toDataURL(safe(d.businessDetails?.bin), { width: 80 }, (err: any, qrDataUrl: string) => {
      const docDef: any = {
        pageSize: 'A4',
        pageMargins: [50, 40, 50, 40],
        defaultStyle: { font: 'PlaywriteCU', fontSize: 10 },
        content: [
          // Top right: Mushak-2.3 box
          {
            columns: [
              { text: '', width: '*' },
              {
                table: {
                  body: [
                    [
                      {
                        text: safe(l.titles?.m_name),
                        bold: true,
                        fontSize: 9,
                        margin: [4, 2, 4, 2],
                      },
                    ],
                  ],
                },
                layout: 'noBorders',
                width: 'auto',
              },
            ],
            margin: [0, 0, 0, 5],
          },

          // Government Header
          {
            stack: [
              { text: safe(l.titles?.gov), alignment: 'center', bold: true, fontSize: 12 },
              {
                text: safe(l.titles?.nbr),
                alignment: 'center',
                bold: true,
                fontSize: 11,
                margin: [0, 2, 0, 8],
              },
              { text: safe(l.titles?.commissionerate), alignment: 'center', fontSize: 9 },
              {
                text: safe(l.titles?.division),
                alignment: 'center',
                fontSize: 9,
                margin: [0, 0, 0, 10],
              },
            ],
          },

          // Certificate Title
          {
            text: safe(l.titles?.form),
            alignment: 'center',
            bold: true,
            fontSize: 13,
            margin: [0, 0, 0, 8],
          },

          // Sub text
          {
            text: safe(l.titles?.rule),
            alignment: 'center',
            fontSize: 9,
            margin: [40, 0, 40, 12],
          },

          // BIN - bold, large, underline
          {
            text: `BIN : ${safe(d.businessDetails?.bin)}`,
            alignment: 'center',
            bold: true,
            fontSize: 15,
            decoration: 'underline',
            margin: [0, 0, 0, 15],
          },

          // Details - label : value format (no table border, just rows)
          {
            table: {
              widths: [160, 10, '*'],
              body: [
                [
                  { text: safe(l.info?.name_of_entity), border: [false, false, false, false] },
                  { text: ':', border: [false, false, false, false] },
                  {
                    text: safe(d.businessDetails?.nameOfEntity),
                    border: [false, false, false, false],
                  },
                ],
                [
                  { text: safe(l.info?.trading_brand_name), border: [false, false, false, false] },
                  { text: ':', border: [false, false, false, false] },
                  {
                    text: safe(d.businessDetails?.tradingBrandName),
                    border: [false, false, false, false],
                  },
                ],
                [
                  { text: safe(l.info?.old_bin), border: [false, false, false, false] },
                  { text: ':', border: [false, false, false, false] },
                  { text: safe(d.businessDetails?.oldBIN), border: [false, false, false, false] },
                ],
                [
                  { text: safe(l.info?.etin), border: [false, false, false, false] },
                  { text: ':', border: [false, false, false, false] },
                  { text: safe(d.businessDetails?.eTIN), border: [false, false, false, false] },
                ],
                [
                  { text: safe(l.info?.address), border: [false, false, false, false] },
                  { text: ':', border: [false, false, false, false] },
                  {
                    text: safe(d.businessDetails?.address?.fullAddress),
                    border: [false, false, false, false],
                  },
                ],
                [
                  { text: safe(l.info?.issue_date), border: [false, false, false, false] },
                  { text: ':', border: [false, false, false, false] },
                  {
                    text: safe(d.registrationInfo?.issueDate),
                    border: [false, false, false, false],
                  },
                ],
                [
                  { text: safe(l.info?.effective_date), border: [false, false, false, false] },
                  { text: ':', border: [false, false, false, false] },
                  {
                    text: safe(d.registrationInfo?.effectiveDate),
                    border: [false, false, false, false],
                  },
                ],
                [
                  { text: safe(l.info?.type_of_ownership), border: [false, false, false, false] },
                  { text: ':', border: [false, false, false, false] },
                  {
                    text: safe(d.registrationInfo?.typeOfOwnership),
                    border: [false, false, false, false],
                  },
                ],
                [
                  { text: safe(l.info?.major_area), border: [false, false, false, false] },
                  { text: ':', border: [false, false, false, false] },
                  {
                    text: safe(d.registrationInfo?.majorAreaOfEconomicActivity),
                    border: [false, false, false, false],
                  },
                ],
              ],
            },
            margin: [0, 0, 0, 30],
          },

          // QR Code center
          {
            image: qrDataUrl,
            width: 80,
            alignment: 'center',
            margin: [0, 0, 0, 20],
          },

          // Footer note
          {
            text: safe(l.footer?.note),
            alignment: 'center',
            fontSize: 8,
            italics: true,
          },
        ],
      };

      pdfMake.createPdf(docDef).download(`Mushak_2.3_${lang}.pdf`);
    });
  }

  exportmushak_6_1_English(data: any, lang: string) {
    const l = (data.labels?.mushak_6_1 || {}) as any;
    const targetData = data.mushak_6_1_data?.[lang] || {};
    const info = (targetData.companyInfo || {}) as any;
    const items = (targetData.items || []) as any[];
    const sh = (data.labels?.mushak_6_1?.sub_headers || {}) as any;

    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : ' ');

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
    };

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
            widths: ['15%', '2%', '83%'],
            body: [
              [l.info?.comp_name, ':', safe(info.name)],
              [l.info?.address, ':', safe(info.address)],
              [l.info?.bin, ':', safe(info.bin)],
            ],
          },
          layout: 'noBorders',
        },

        // Main Table (21 Columns as per PDF)
        {
          table: {
            headerRows: 4,
            widths: [
              15, 33, 18, 25, 35, 33, 45, 40, 33, 45, 22, 25, 22, 22, 20, 25, 20, 25, 24, 22, 24,
            ],
            body: [
              // Row 1: Merged Headers
              [
                {
                  text: safe(l.titles?.sub_title),
                  colSpan: 21,
                  alignment: 'center',
                  bold: true,
                  margin: [0, 2, 0, 2],
                },
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
              ],

              [
                { text: l.headers?.sl, rowSpan: 3, bold: true, alignment: 'center' },
                { text: l.headers?.date, rowSpan: 3, bold: true, alignment: 'center' },
                { text: l.headers?.opening_stock, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: l.headers?.purchase_info, colSpan: 14, alignment: 'center', bold: true },
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                { text: l.headers?.closing_stock, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: l.headers?.remarks, rowSpan: 3, bold: true, alignment: 'center' },
              ],

              [
                {},
                {},
                sh.qty || ' ',
                sh.val || ' ',
                { text: l.headers?.invoice_info, rowSpan: 2, alignment: 'center', bold: true },
                { text: sh.date || ' ', rowSpan: 2, alignment: 'center', bold: true },
                { text: l.headers?.seller_info, colSpan: 3, alignment: 'center', bold: true },
                {},
                {},
                { text: l.headers?.item_desc, rowSpan: 2, bold: true, alignment: 'center' },
                { text: sh.qty || ' ', rowSpan: 2, bold: true, alignment: 'center' },
                { text: sh.val || ' ', rowSpan: 2, bold: true, alignment: 'center' },
                { text: sh.sd || ' ', rowSpan: 2, bold: true, alignment: 'center' },
                { text: sh.vat || ' ', rowSpan: 2, bold: true, alignment: 'center' },
                { text: l.headers?.total_materials, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: l.headers?.prod_info, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: l.headers?.qty || ' ', rowSpan: 2, alignment: 'center', bold: true },
                { text: l.headers?.value || ' ', rowSpan: 2, alignment: 'center', bold: true },
                {},
              ],
              // Row 2: Sub-headers
              [
                {},
                {},
                {},
                {},
                sh.no || ' ',
                {},
                sh.name || ' ',
                sh.addr || ' ',
                sh.bin || ' ',
                {},
                {},
                {},
                {},
                {},
                sh.qty || ' ',
                sh.val || ' ',
                sh.qty || ' ',
                sh.val || ' ',
                {},
                {},
                {},
              ],
              // Row 3: Column Numbers (1) to (21)
              Array.from({ length: 21 }, (_, i) => ({
                text: `(${i + 1})`,
                alignment: 'center',
                fontSize: 5,
              })),

              // Data Rows from db.json
              ...items.map((item: any) => [
                safe(item.sl),
                safe(item.date),
                safe(item.opening_qty),
                safe(item.opening_val),
                safe(item.invoice_no),
                safe(item.invoice_date),
                safe(item.seller_name),
                safe(item.seller_address),
                safe(item.seller_bin),
                safe(item.item_desc),
                safe(item.purchase_qty),
                safe(item.purchase_val),
                safe(item.sd),
                safe(item.vat),
                safe(item.total_qty),
                safe(item.total_val),
                safe(item.used_qty),
                safe(item.used_val),
                safe(item.closing_qty),
                safe(item.closing_val),
                safe(item.remarks),
              ]),
            ],
          },
        },
        {
          margin: [0, 20, 0, 0],
          stack: [
            {
              text: l.footer?.special_note_title,
              bold: true,
              decoration: 'underline',
              fontSize: 8,
            },
            {
              ul: (l.footer?.notes || []).map((note: string) => ({
                text: note,
                margin: [0, 2, 0, 0],
              })),
              fontSize: 7,
              lineHeight: 1.3,
            },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`mushak_6_1_${lang}.pdf`);
  }

  exportmushak_6_1_Bangla(data: any, lang: string) {
    const l = (data.labels?.mushak_6_1 || {}) as any;
    const targetData = data.mushak_6_1_data?.[lang] || {};
    const info = (targetData.companyInfo || {}) as any;
    const items = (targetData.items || []) as any[];
    const sh = (data.labels?.mushak_6_1?.sub_headers || {}) as any;

    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : ' ');

    (pdfMake as any).fonts = {
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
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
            widths: ['15%', '2%', '83%'],
            body: [
              [l.info?.comp_name, ':', safe(info.name)],
              [l.info?.address, ':', safe(info.address)],
              [l.info?.bin, ':', safe(info.bin)],
            ],
          },
          layout: 'noBorders',
        },

        {
          table: {
            headerRows: 5,
            // widths: [15, 32, 17, 25, 30, 32, 40, 40, 30, 40, 17, 25, 20, 20, 17, 25, 17, 25, 20, 20, 20],
            widths: [
              15, 32, 17, 28, 35, 32, 45, 40, 30, 45, 17, 28, 28, 28, 17, 25, 17, 25, 20, 20, 20,
            ],
            body: [
              // Row 1: Merged Headers
              [
                {
                  text: safe(l.titles?.sub_title),
                  colSpan: 21,
                  alignment: 'center',
                  bold: true,
                  margin: [0, 2, 0, 2],
                },
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
              ],

              [
                { text: l.headers?.sl, rowSpan: 3, bold: true, alignment: 'center' },
                { text: l.headers?.date, rowSpan: 3, bold: true, alignment: 'center' },
                { text: l.headers?.opening_stock, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: l.headers?.purchase_info, colSpan: 14, alignment: 'center', bold: true },
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                { text: l.headers?.closing_stock, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: l.headers?.remarks, rowSpan: 3, bold: true, alignment: 'center' },
              ],

              [
                {},
                {},
                sh.qty || ' ',
                sh.val || ' ',
                { text: l.headers?.invoice_info, rowSpan: 2, alignment: 'center', bold: true },
                { text: sh.date || ' ', rowSpan: 2, alignment: 'center', bold: true },
                { text: l.headers?.seller_info, colSpan: 3, alignment: 'center', bold: true },
                {},
                {},
                { text: l.headers?.item_desc, rowSpan: 2, bold: true, alignment: 'center' },
                { text: sh.qty || ' ', rowSpan: 2, bold: true, alignment: 'center' },
                { text: sh.val || ' ', rowSpan: 2, bold: true, alignment: 'center' },
                { text: sh.sd || ' ', rowSpan: 2, bold: true, alignment: 'center' },
                { text: sh.vat || ' ', rowSpan: 2, bold: true, alignment: 'center' },
                { text: l.headers?.total_materials, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: l.headers?.prod_info, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: l.headers?.qty || ' ', rowSpan: 2, alignment: 'center', bold: true },
                { text: l.headers?.value || ' ', rowSpan: 2, alignment: 'center', bold: true },
                {},
              ],
              // Row 2: Sub-headers
              [
                {},
                {},
                {},
                {},
                sh.no || ' ',
                {},
                sh.name || ' ',
                sh.addr || ' ',
                sh.bin || ' ',
                {},
                {},
                {},
                {},
                {},
                sh.qty || ' ',
                sh.val || ' ',
                sh.qty || ' ',
                sh.val || ' ',
                {},
                {},
                {},
              ],
              // Row 3: Column Numbers (1) to (21)
              // Array.from({ length: 21 }, (_, i) => ({ text: `(${i + 1})`, alignment: 'center', fontSize: 5 })),
              Array.from({ length: 21 }, (_, i) => {
                const banglaDigits = [
                  '০',
                  '১',
                  '২',
                  '৩',
                  '৪',
                  '৫',
                  '৬',
                  '৭',
                  '৮',
                  '৯',
                  '১১',
                  '১২',
                  '১৩',
                  '১৪',
                  '১৫',
                  '১৬',
                  '১৭',
                  '১৮',
                  '১৯',
                  '২০',
                  '২১',
                ];
                const englishNumber = (i + 1).toString();
                const bNum = englishNumber
                  .split('')
                  .map((d) => banglaDigits[parseInt(d)])
                  .join('');
                return {
                  text: `(${bNum})`,
                  alignment: 'center',
                  fontSize: 4.5,
                  fillColor: '#f5f5f5',
                };
              }),

              // Data Rows from db.json
              ...items.map((item: any) => [
                safe(item.sl),
                safe(item.date),
                safe(item.opening_qty),
                safe(item.opening_val),
                safe(item.invoice_no),
                safe(item.invoice_date),
                safe(item.seller_name),
                safe(item.seller_address),
                safe(item.seller_bin),
                safe(item.item_desc),
                safe(item.purchase_qty),
                safe(item.purchase_val),
                safe(item.sd),
                safe(item.vat),
                safe(item.total_qty),
                safe(item.total_val),
                safe(item.used_qty),
                safe(item.used_val),
                safe(item.closing_qty),
                safe(item.closing_val),
                safe(item.remarks),
              ]),
            ],
          },
        },

        {
          margin: [0, 20, 0, 0],
          stack: [
            {
              text: l.footer?.special_note_title,
              bold: true,
              decoration: 'underline',
              fontSize: 8,
            },
            {
              ul: (l.footer?.notes || []).map((note: string) => ({
                text: note,
                margin: [0, 2, 0, 0],
              })),
              fontSize: 7,
              lineHeight: 1.3,
            },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`mushak_6_1_${lang}.pdf`);
  }

  exportMushak_6_2_English(data: any, lang: string) {
    const labels = (data.labels?.mushak_6_2 || {}) as any;
    const targetData = data.mushak_6_2_data?.[lang] || {};
    const info = (targetData.companyInfo || {}) as any;
    const items = (targetData.items || []) as any[];
    const sh = (data.labels?.mushak_6_2?.sub_headers || {}) as any;

    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : ' ');

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageOrientation: 'landscape',
      defaultStyle: { font: 'Nunito', fontSize: 5.5 },
      content: [
        { text: safe(labels.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(labels.titles?.gov), alignment: 'center', bold: true },
        { text: safe(labels.titles?.nbr), alignment: 'center', bold: true },
        { text: safe(labels.titles?.form), alignment: 'center', bold: true, fontSize: 10 },
        { text: safe(labels.titles?.rule), alignment: 'center', fontSize: 7, margin: [0, 0, 0, 5] },

        {
          margin: [0, 10, 0, 0],
          table: {
            headerRows: 4,
            widths: [
              15, 33, 22, 25, 22, 33, 22, 30, 45, 45, 30, 30, 25, 45, 22, 25, 25, 30, 24, 22, 20,
            ],
            body: [
              [
                {
                  text: safe(labels.titles?.sub_title),
                  colSpan: 21,
                  alignment: 'center',
                  bold: true,
                  margin: [0, 2, 0, 2],
                },
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
              ],

              [
                { text: labels.headers?.sl, rowSpan: 2, bold: true, alignment: 'center' },
                { text: labels.headers?.date, rowSpan: 2, bold: true, alignment: 'center' },
                { text: labels.headers?.opening, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: labels.headers?.production, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: labels.headers?.total, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: labels.headers?.buyer, colSpan: 3, alignment: 'center', bold: true },
                {},
                {},
                { text: labels.headers?.invoice, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: labels.headers?.sales, colSpan: 5, alignment: 'center', bold: true },
                {},
                {},
                {},
                {},
                { text: labels.headers?.closing, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: labels.headers?.remarks, rowSpan: 2, bold: true, alignment: 'center' },
              ],

              [
                {},
                {},
                safe(sh.qty),
                safe(sh.val),
                safe(sh.qty),
                safe(sh.val),
                safe(sh.qty),
                safe(sh.val),
                safe(sh.name),
                safe(sh.addr),
                safe(sh.bin),
                safe(sh.date),
                safe(sh.no),
                safe(sh.item_desc),
                safe(sh.qty),
                safe(sh.tax),
                safe(sh.sd),
                safe(sh.vat),
                safe(sh.qty),
                safe(sh.val),
                {},
              ],

              Array.from({ length: 21 }, (_, i) => ({
                text: `(${i + 1})`,
                alignment: 'center',
                fontSize: 4.5,
                fillColor: '#f5f5f5',
              })),

              ...items.map((item: any) => [
                safe(item.sl),
                safe(item.date),
                safe(item.opening_qty),
                safe(item.opening_val),
                safe(item.prod_qty),
                safe(item.prod_val),
                safe(item.total_qty),
                safe(item.total_val),
                safe(item.buyer_name),
                safe(item.buyer_address),
                safe(item.buyer_bin),
                safe(item.invoice_date),
                safe(item.invoice_no),
                safe(item.item_desc),
                safe(item.sales_qty),
                safe(item.sales_val),
                safe(item.sd),
                safe(item.vat),
                safe(item.closing_qty),
                safe(item.closing_val),
                safe(item.remarks),
              ]),
            ],
          },
        },
        {
          margin: [0, 15, 0, 0],
          stack: [
            {
              text: safe(labels.footer?.note_title),
              bold: true,
              decoration: 'underline',
              fontSize: 8,
            },
            {
              ol: (labels.footer?.notes || []).map((note: string) => ({
                text: note,
                margin: [0, 2, 0, 0],
              })),
              fontSize: 6.5,
            },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.2_${lang}.pdf`);
  }

  exportMushak_6_2_Bangla(data: any, lang: string) {
    const labels = (data.labels?.mushak_6_2 || {}) as any;
    const targetData = data.mushak_6_2_data?.[lang] || {};
    const info = (targetData.companyInfo || {}) as any;
    const items = (targetData.items || []) as any[];
    const sh = (data.labels?.mushak_6_2?.sub_headers || {}) as any;

    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : ' ');

    (pdfMake as any).fonts = {
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageOrientation: 'landscape',
      defaultStyle: { font: 'PlaywriteCU', fontSize: 5.5 },
      content: [
        { text: safe(labels.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(labels.titles?.gov), alignment: 'center', bold: true },
        { text: safe(labels.titles?.nbr), alignment: 'center', bold: true },
        { text: safe(labels.titles?.form), alignment: 'center', bold: true, fontSize: 10 },
        { text: safe(labels.titles?.rule), alignment: 'center', fontSize: 7, margin: [0, 0, 0, 5] },

        {
          margin: [0, 10, 0, 0],
          table: {
            headerRows: 4,
            widths: [
              15, 30, 16, 28, 16, 28, 16, 28, 45, 45, 25, 30, 30, 45, 16, 25, 25, 25, 16, 28, 20,
            ],
            body: [
              [
                {
                  text: safe(labels.titles?.sub_title),
                  colSpan: 21,
                  alignment: 'center',
                  bold: true,
                  margin: [0, 2, 0, 2],
                },
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
              ],

              [
                { text: labels.headers?.sl, rowSpan: 2, bold: true, alignment: 'center' },
                { text: labels.headers?.date, rowSpan: 2, bold: true, alignment: 'center' },
                { text: labels.headers?.opening, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: labels.headers?.production, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: labels.headers?.total, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: labels.headers?.buyer, colSpan: 3, alignment: 'center', bold: true },
                {},
                {},
                { text: labels.headers?.invoice, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: labels.headers?.sales, colSpan: 5, alignment: 'center', bold: true },
                {},
                {},
                {},
                {},
                { text: labels.headers?.closing, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: labels.headers?.remarks, rowSpan: 2, bold: true, alignment: 'center' },
              ],

              [
                {},
                {},
                sh.qty,
                sh.val,
                sh.qty,
                sh.val,
                sh.qty,
                sh.val,
                sh.name || ' ',
                sh.addr || ' ',
                sh.bin || ' ',
                sh.date || ' ',
                sh.no || ' ',
                sh.item_desc || ' ',
                sh.qty,
                sh.tax || ' ',
                sh.sd,
                sh.vat,
                sh.qty,
                sh.val,
                {},
              ],

              // Array.from({ length: 21 }, (_, i) => ({ text: `(${i + 1})`, alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' })),
              // Array.from({ length: 21 }, (_, i) => {
              //   const banglaDigits = ['০', '১', '২', '৩', '৪', '৫', '৬', '৭=৩+৪', '৮=৪+৬', '৯', '১১', '১২', '১৩', '১৪', '১৫', '১৬', '১৭', '১৮', '১৯', '২০', '২১'];
              //   const englishNumber = (i + 1).toString();
              //   const bNum = englishNumber.split('').map(d => banglaDigits[parseInt(d)]).join('');
              //   return { text: `(${bNum})`, alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' };
              // }),
              [
                { text: '(১)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
                { text: '(২)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
                { text: '(৩)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
                { text: '(৪)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
                { text: '(৫)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
                { text: '(৬)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
                { text: '(৭)\n=(৩+৫)', alignment: 'center', fontSize: 4, fillColor: '#f5f5f5' },
                { text: '(৮)\n=(৪+৬)', alignment: 'center', fontSize: 4, fillColor: '#f5f5f5' },
                { text: '(৯)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
                { text: '(১০)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
                { text: '(১১)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
                { text: '(১২)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
                { text: '(১৩)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
                { text: '(১৪)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
                { text: '(১৫)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
                { text: '(১৬)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
                { text: '(১৭)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
                { text: '(১৮)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
                { text: '(১৯)\n=(৭-১৫)', alignment: 'center', fontSize: 4, fillColor: '#f5f5f5' },
                { text: '(২০)\n=(৮-১৬)', alignment: 'center', fontSize: 4, fillColor: '#f5f5f5' },
                { text: '(২১)', alignment: 'center', fontSize: 4.5, fillColor: '#f5f5f5' },
              ],

              ...items.map((item: any) => [
                safe(item.sl),
                safe(item.date),
                safe(item.opening_qty),
                safe(item.opening_val),
                safe(item.prod_qty),
                safe(item.prod_val),
                safe(item.total_qty),
                safe(item.total_val),
                safe(item.buyer_name),
                safe(item.buyer_address),
                safe(item.buyer_bin),
                safe(item.invoice_date),
                safe(item.invoice_no),
                safe(item.item_desc),
                safe(item.sales_qty),
                safe(item.sales_val),
                safe(item.sd),
                safe(item.vat),
                safe(item.closing_qty),
                safe(item.closing_val),
                safe(item.remarks),
              ]),
            ],
          },
        },
        {
          margin: [0, 15, 0, 0],
          stack: [
            {
              text: safe(labels.footer?.note_title),
              bold: true,
              decoration: 'underline',
              fontSize: 8,
            },
            {
              ul: (labels.footer?.notes || []).map((note: string) => ({
                text: note,
                margin: [0, 2, 0, 0],
              })),
              fontSize: 6.5,
            },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.2_${lang}.pdf`);
  }

  exportMushak_6_2_1_English(data: any, lang: string) {
    const l = (data.labels?.mushak_6_2_1 || {}) as any;
    const targetData = data.mushak_6_2_1_data?.[lang] || {};
    const items = (targetData.items || []) as any[];
    const sh = (data.labels?.mushak_6_2_1?.sub_headers || {}) as any;

    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : ' ');

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageOrientation: 'landscape',
      pageMargins: [18, 30, 18, 30],
      defaultStyle: { font: 'Nunito', fontSize: 5 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        { text: safe(l.titles?.form), alignment: 'center', bold: true, fontSize: 10 },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 6.5, margin: [0, 0, 0, 5] },

        {
          table: {
            headerRows: 4,
            widths: [
              11, 27, 20, 20, 20, 20, 20, 20, 37, 35, 25, 19, 27, 37, 20, 20, 17, 20, 25, 25, 20,
              19, 20, 20, 20, 15,
            ],
            body: [
              [
                { text: l.headers?.sl, rowSpan: 3, bold: true },
                { text: l.headers?.date, rowSpan: 3, bold: true },
                { text: l.headers?.opening, colSpan: 2, bold: true, alignment: 'center' },
                {},
                { text: l.headers?.purchase, colSpan: 2, bold: true, alignment: 'center' },
                {},
                { text: l.headers?.total, colSpan: 2, bold: true, alignment: 'center' },
                {},
                { text: l.headers?.seller, colSpan: 3, bold: true, alignment: 'center' },
                {},
                {},
                { text: l.headers?.p_invoice, colSpan: 2, bold: true, alignment: 'center' },
                {},
                { text: l.headers?.sales_desc, colSpan: 5, bold: true, alignment: 'center' },
                {},
                {},
                {},
                {},
                { text: l.headers?.buyer, colSpan: 3, bold: true, alignment: 'center' },
                {},
                {},
                { text: l.headers?.s_invoice, colSpan: 2, bold: true, alignment: 'center' },
                {},
                { text: l.headers?.closing, colSpan: 2, bold: true, alignment: 'center' },
                {},
                { text: l.headers?.remarks, rowSpan: 3, bold: true },
              ],
              [
                {},
                {},
                { text: sh.qty, rowSpan: 2, bold: true },
                { text: sh.val, rowSpan: 2, bold: true },
                { text: sh.qty, rowSpan: 2, bold: true },
                { text: sh.val, rowSpan: 2, bold: true },
                { text: sh.qty, rowSpan: 2, bold: true },
                { text: sh.val, rowSpan: 2, bold: true },
                { text: '', colSpan: 3 },
                {},
                {},
                sh.no,
                sh.date,
                sh.item_desc,
                sh.qty,
                sh.tax,
                sh.sd,
                sh.vat,
                { text: '', colSpan: 3 },
                {},
                {},
                { text: '', colSpan: 2 },
                {},
                sh.qty,
                sh.val,
                {},
              ],
              [
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                sh.name,
                sh.addr,
                sh.bin,
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                sh.name,
                sh.addr,
                sh.bin,
                sh.no,
                sh.date,
                {},
                {},
                {},
              ],
              [
                '(1)',
                '(2)',
                '(3)',
                '(4)',
                '(5)',
                '(6)',
                { text: '(7)\n=(3+5)', fontSize: 4 },
                { text: '(8)\n=(4+6)', fontSize: 4 },
                '(9)',
                '(10)',
                '(11)',
                '(12)',
                '(13)',
                '(14)',
                '(15)',
                '(16)',
                '(17)',
                '(18)',
                '(19)',
                '(20)',
                '(21)',
                '(22)',
                '(23)',
                { text: '(24)\n=(7-15)', fontSize: 4 },
                { text: '(25)\n=(8-16)', fontSize: 4 },
                '(26)',
              ].map((num) => ({
                text: typeof num === 'string' ? num : num.text,
                alignment: 'center',
                fillColor: '#f5f5f5',
                fontSize: 4.5,
              })),

              // Data Rows
              ...items.map((item) => [
                safe(item.sl),
                safe(item.date),
                safe(item.op_qty),
                safe(item.op_val),
                safe(item.p_qty),
                safe(item.p_val),
                safe(item.tot_qty),
                safe(item.tot_val),
                safe(item.s_name),
                safe(item.s_addr),
                safe(item.s_bin),
                safe(item.p_inv_no),
                safe(item.p_inv_date),
                safe(item.item_desc),
                safe(item.s_qty),
                safe(item.s_val),
                safe(item.sd),
                safe(item.vat),
                safe(item.b_name),
                safe(item.b_addr),
                safe(item.b_bin),
                safe(item.s_inv_no),
                safe(item.s_inv_date),
                safe(item.cl_qty),
                safe(item.cl_val),
                safe(item.remarks),
              ]),
            ],
          },
        },

        // Footer Notes [cite: 17, 18, 19]
        {
          margin: [0, 15, 0, 0],
          stack: [
            { text: safe(l.footer?.note_title), bold: true, decoration: 'underline', fontSize: 8 },
            {
              ul: (l.footer?.notes || []).map((note: string) => ({
                text: note,
                margin: [0, 2, 0, 0],
              })),
              fontSize: 6.5,
            },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.2.1_${lang}.pdf`);
  }

  exportMushak_6_2_1_Bangla(data: any, lang: string) {
    const l = (data.labels?.mushak_6_2_1 || {}) as any;
    const targetData = data.mushak_6_2_1_data?.[lang] || {};
    const items = (targetData.items || []) as any[];
    const sh = (data.labels?.mushak_6_2_1?.sub_headers || {}) as any;

    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : ' ');

    (pdfMake as any).fonts = {
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageOrientation: 'landscape',
      pageMargins: [20, 30, 20, 30],
      defaultStyle: { font: 'PlaywriteCU', fontSize: 5 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        { text: safe(l.titles?.form), alignment: 'center', bold: true, fontSize: 10 },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 6.5, margin: [0, 0, 0, 5] },

        {
          table: {
            headerRows: 4,
            widths: [
              12, 26, 15, 22, 15, 22, 15, 22, 40, 35, 25, 20, 26, 40, 15, 22, 16, 20, 25, 25, 20,
              20, 20, 18, 20, 12,
            ],
            body: [
              [
                { text: l.headers?.sl, rowSpan: 3, bold: true },
                { text: l.headers?.date, rowSpan: 3, bold: true },
                { text: l.headers?.opening, colSpan: 2, bold: true, alignment: 'center' },
                {},
                { text: l.headers?.purchase, colSpan: 2, bold: true, alignment: 'center' },
                {},
                { text: l.headers?.total, colSpan: 2, bold: true, alignment: 'center' },
                {},
                { text: l.headers?.seller, colSpan: 3, bold: true, alignment: 'center' },
                {},
                {},
                { text: l.headers?.p_invoice, colSpan: 2, bold: true, alignment: 'center' },
                {},
                { text: l.headers?.sales_desc, colSpan: 5, bold: true, alignment: 'center' },
                {},
                {},
                {},
                {},
                { text: l.headers?.buyer, colSpan: 3, bold: true, alignment: 'center' },
                {},
                {},
                { text: l.headers?.s_invoice, colSpan: 2, bold: true, alignment: 'center' },
                {},
                { text: l.headers?.closing, colSpan: 2, bold: true, alignment: 'center' },
                {},
                { text: l.headers?.remarks, rowSpan: 3, bold: true },
              ],
              [
                {},
                {},
                { text: sh.qty, rowSpan: 2, bold: true },
                { text: sh.val, rowSpan: 2, bold: true },
                { text: sh.qty, rowSpan: 2, bold: true },
                { text: sh.val, rowSpan: 2, bold: true },
                { text: sh.qty, rowSpan: 2, bold: true },
                { text: sh.val, rowSpan: 2, bold: true },
                { text: '', colSpan: 3 },
                {},
                {},
                sh.no,
                sh.date,
                sh.item_desc,
                sh.qty,
                sh.tax,
                sh.sd,
                sh.vat,
                { text: '', colSpan: 3 },
                {},
                {},
                { text: '', colSpan: 2 },
                {},
                sh.qty,
                sh.val,
                {},
              ],
              [
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                sh.name,
                sh.addr,
                sh.bin,
                {},
                {},
                {},
                {},
                {},
                {},
                {},
                sh.name,
                sh.addr,
                sh.bin,
                sh.no,
                sh.date,
                {},
                {},
                {},
              ],
              [
                '(১)',
                '(২)',
                '(৩)',
                '(৪)',
                '(৫)',
                '(৬)',
                { text: '(৭)\n=(৩+৫)', fontSize: 4 },
                { text: '(৮)\n=(৪+৬)', fontSize: 4 },
                '(৯)',
                '(১০)',
                '(১১)',
                '(১২)',
                '(১৩)',
                '(১৪)',
                '(১৫)',
                '(১৬)',
                '(১৭)',
                '(১৮)',
                '(১৯)',
                '(২০)',
                '(২১)',
                '(২২)',
                '(২৩)',
                { text: '(২৪)\n=(৭-১৫)', fontSize: 4 },
                { text: '(২৫)\n=(৮-১৬)', fontSize: 4 },
                '(২৬)',
              ].map((num) => ({
                text: typeof num === 'string' ? num : num.text,
                alignment: 'center',
                fillColor: '#f5f5f5',
                fontSize: 4.5,
              })),

              // Data Rows
              ...items.map((item) => [
                safe(item.sl),
                safe(item.date),
                safe(item.op_qty),
                safe(item.op_val),
                safe(item.p_qty),
                safe(item.p_val),
                safe(item.tot_qty),
                safe(item.tot_val),
                safe(item.s_name),
                safe(item.s_addr),
                safe(item.s_bin),
                safe(item.p_inv_no),
                safe(item.p_inv_date),
                safe(item.item_desc),
                safe(item.s_qty),
                safe(item.s_val),
                safe(item.sd),
                safe(item.vat),
                safe(item.b_name),
                safe(item.b_addr),
                safe(item.b_bin),
                safe(item.s_inv_no),
                safe(item.s_inv_date),
                safe(item.cl_qty),
                safe(item.cl_val),
                safe(item.remarks),
              ]),
            ],
          },
        },

        // Footer Notes [cite: 17, 18, 19]
        {
          margin: [0, 15, 0, 0],
          stack: [
            { text: safe(l.footer?.note_title), bold: true, decoration: 'underline', fontSize: 8 },
            {
              ul: (l.footer?.notes || []).map((note: string) => ({
                text: note,
                margin: [0, 2, 0, 0],
              })),
              fontSize: 6.5,
            },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.2.1_${lang}.pdf`);
  }

  exportMushak_6_3_English(data: any, lang: string) {
    const l = (data.labels?.mushak_6_3 || {}) as any;
    const targetData = data.mushak_6_3_data?.[lang] || {};
    const items = (targetData.items || []) as any[];

    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: lang === 'BN' ? 'PlaywriteCU' : 'Nunito', fontSize: 8 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        { text: safe(l.titles?.form), alignment: 'center', bold: true, fontSize: 12 },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 8, margin: [0, 0, 0, 15] },

        {
          columns: [
            {
              width: '*',
              stack: [
                { text: `${l.info?.buyer_name}: ${safe(targetData.buyer_name)}` },
                { text: `${l.info?.buyer_bin}: ${safe(targetData.buyer_bin)}` },
                { text: `${l.info?.dest_addr}: ${safe(targetData.delivery_dest)}` },
              ],
            },
            {
              width: '25%',
              stack: [
                { text: `${l.info?.inv_no}:`, bold: true },
                { text: `${l.info?.inv_date}:` },
                { text: `${l.info?.inv_time}:` },
              ],
              alignment: 'right',
              margin: [0, 0, 10, 0],
            },
            {
              width: '20%',
              stack: [
                { text: safe(targetData.invoice_no), bold: true },
                { text: safe(targetData.issue_date) },
                { text: safe(targetData.issue_time) },
              ],
              alignment: 'left',
            },
          ],
          margin: [0, 0, 0, 10],
        },

        {
          table: {
            headerRows: 1,
            widths: [20, '*', 35, 35, 45, 45, 40, 45, 45, 50],
            body: [
              [
                { text: l.headers?.sl, bold: true, alignment: 'center' },
                { text: l.headers?.desc, bold: true, alignment: 'center' },
                { text: l.headers?.unit, bold: true, alignment: 'center' },
                { text: l.headers?.qty, bold: true, alignment: 'center' },
                { text: l.headers?.u_price, bold: true, alignment: 'center' },
                { text: l.headers?.t_price, bold: true, alignment: 'center' },
                { text: l.headers?.sd_amount, bold: true, alignment: 'center' },
                { text: l.headers?.vat_rate, bold: true, alignment: 'center' },
                { text: l.headers?.vat_amount, bold: true, alignment: 'center' },
                { text: l.headers?.total_incl_all, bold: true, alignment: 'center' },
              ],
              ...items.map((item, index) => [
                { text: (index + 1).toString(), alignment: 'center' },
                safe(item.desc),
                safe(item.unit),
                safe(item.qty),
                safe(item.u_price),
                safe(item.t_price),
                safe(item.sd_amount),
                safe(item.vat_rate),
                safe(item.vat_amount),
                safe(item.total_incl_all),
              ]),
              [
                {
                  text: l.headers?.grand_total || 'সর্বমোট',
                  colSpan: 5,
                  alignment: 'right',
                  bold: true,
                },
                {},
                {},
                {},
                {},
                { text: safe(targetData.total_t_price), bold: true },
                { text: safe(targetData.total_sd), bold: true },
                {},
                { text: safe(targetData.total_vat), bold: true },
                { text: safe(targetData.grand_total), bold: true },
              ],
            ],
          },
        },
        {
          margin: [0, 15, 0, 0],
          stack: [
            {
              width: 'auto',
              stack: [
                { text: `${l.footer?.auth_label}: ${safe(targetData.auth_person)}`, bold: true },
                { text: `${l.footer?.designation_label}: ${safe(targetData.designation)}` },
                { text: `${l.footer?.signature_label}: ${safe(targetData.buyer_signature)}` },
                { text: `${l.footer?.seal_label}: ${safe(targetData.buyer_seal)}` },
              ],
              alignment: 'left',
            },
            {
              margin: [0, 10, 0, 0],
              text: typeof l.footer?.note === 'string' ? l.footer.note : '',
              fontSize: 6.5,
            },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.3_${lang}.pdf`);
  }

  exportMushak_6_3_Bangla(data: any, lang: string) {
    const l = (data.labels?.mushak_6_3 || {}) as any;
    const targetData = data.mushak_6_3_data?.[lang] || {};
    const items = (targetData.items || []) as any[];

    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');

    (pdfMake as any).fonts = {
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
    };
    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: lang === 'BN' ? 'PlaywriteCU' : 'Nunito', fontSize: 8 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        { text: safe(l.titles?.form), alignment: 'center', bold: true, fontSize: 12 },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 8, margin: [0, 0, 0, 15] },

        {
          columns: [
            {
              width: '*',
              stack: [
                { text: `${l.info?.buyer_name}: ${safe(targetData.buyer_name)}` },
                { text: `${l.info?.buyer_bin}: ${safe(targetData.buyer_bin)}` },
                { text: `${l.info?.dest_addr}: ${safe(targetData.delivery_dest)}` },
              ],
            },
            {
              width: '25%',
              stack: [
                { text: `${l.info?.inv_no}:`, bold: true },
                { text: `${l.info?.inv_date}:` },
                { text: `${l.info?.inv_time}:` },
              ],
              alignment: 'right',
              margin: [0, 0, 10, 0],
            },
            {
              width: '20%',
              stack: [
                { text: safe(targetData.invoice_no), bold: true },
                { text: safe(targetData.issue_date) },
                { text: safe(targetData.issue_time) },
              ],
              alignment: 'left',
            },
          ],
          margin: [0, 0, 0, 10],
        },

        {
          table: {
            headerRows: 1,
            widths: [20, '*', 35, 35, 45, 45, 40, 45, 45, 50],
            body: [
              [
                { text: l.headers?.sl, bold: true, alignment: 'center' },
                { text: l.headers?.desc, bold: true, alignment: 'center' },
                { text: l.headers?.unit, bold: true, alignment: 'center' },
                { text: l.headers?.qty, bold: true, alignment: 'center' },
                { text: l.headers?.u_price, bold: true, alignment: 'center' },
                { text: l.headers?.t_price, bold: true, alignment: 'center' },
                { text: l.headers?.sd_amount, bold: true, alignment: 'center' },
                { text: l.headers?.vat_rate, bold: true, alignment: 'center' },
                { text: l.headers?.vat_amount, bold: true, alignment: 'center' },
                { text: l.headers?.total_incl_all, bold: true, alignment: 'center' },
              ],
              ...items.map((item, index) => [
                { text: (index + 1).toString(), alignment: 'center' },
                safe(item.desc),
                safe(item.unit),
                safe(item.qty),
                safe(item.u_price),
                safe(item.t_price),
                safe(item.sd_amount),
                safe(item.vat_rate),
                safe(item.vat_amount),
                safe(item.total_incl_all),
              ]),
              [
                {
                  text: l.headers?.grand_total || 'সর্বমোট',
                  colSpan: 5,
                  alignment: 'right',
                  bold: true,
                },
                {},
                {},
                {},
                {},
                { text: safe(targetData.total_t_price), bold: true },
                { text: safe(targetData.total_sd), bold: true },
                {},
                { text: safe(targetData.total_vat), bold: true },
                { text: safe(targetData.grand_total), bold: true },
              ],
            ],
          },
        },
        {
          margin: [0, 15, 0, 0],
          stack: [
            {
              width: 'auto',
              stack: [
                { text: `${l.footer?.auth_label}: ${safe(targetData.auth_person)}`, bold: true },
                { text: `${l.footer?.designation_label}: ${safe(targetData.designation)}` },
                { text: `${l.footer?.signature_label}: ${safe(targetData.buyer_signature)}` },
                { text: `${l.footer?.seal_label}: ${safe(targetData.buyer_seal)}` },
              ],
              alignment: 'left',
            },
            {
              margin: [0, 10, 0, 0],
              text: typeof l.footer?.note === 'string' ? l.footer.note : '',
              fontSize: 6.5,
            },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.3_${lang}.pdf`);
  }

  exportMushak_6_4_English(data: any, lang: string) {
    const l = (data.labels?.mushak_6_4 || {}) as any;
    const targetData = data.mushak_6_4_data?.[lang] || {};
    const items = (targetData.tableData?.rows || []) as any[];

    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');

    const toBnNum = (n: any) => {
      if (!n) return '';
      return n.toString().replace(/\d/g, (d: any) => '০১২৩৪৫৬৭৮৯'[d]);
    };

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: 'Nunito', fontSize: 9 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        {
          text: safe(l.titles?.form),
          alignment: 'center',
          bold: true,
          fontSize: 13,
          margin: [0, 5, 0, 0],
        },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 8.5, margin: [0, 0, 0, 15] },

        {
          columns: [
            {
              width: '*',
              stack: [
                {
                  text: `${l.info?.registered_person} : ${safe(targetData.formData?.registered_person)}`,
                },
                {
                  text: `${l.info?.registered_person_bin} : ${safe(targetData.formData?.registered_person_bin)}`,
                },
                {
                  text: `${l.info?.chalan_address} : ${safe(targetData.formData?.chalan_address)}`,
                },
              ],
            },
            {
              width: '25%',
              stack: [
                { text: `${l.info?.inv_no} :`, bold: true },
                { text: `${l.info?.inv_date} :` },
                { text: `${l.info?.inv_time} :` },
              ],
              alignment: 'right',
              margin: [0, 0, 10, 0],
            },
            {
              width: '20%',
              stack: [
                { text: safe(targetData.formData?.chalan_number), bold: true },
                { text: safe(targetData.formData?.issue_date) },
                { text: safe(targetData.formData?.issue_time) },
              ],
              alignment: 'left',
            },
          ],
          margin: [0, 0, 0, 10],
        },

        {
          stack: [
            { text: `${l.info?.buyer_name} : ${safe(targetData.formData?.recipient_name)}` },
            { text: `${l.info?.buyer_bin} : ${safe(targetData.formData?.recipient_bin || '-')}` },
            { text: `${l.info?.dest_addr} : ${safe(targetData.formData?.destination)}` },
          ],
          margin: [0, 0, 0, 15],
        },

        {
          table: {
            headerRows: 2,
            widths: [30, '*', '*', 70, 70],
            body: [
              [
                { text: l.headers?.sl, bold: true, alignment: 'center' },
                { text: l.headers?.desc, bold: true, alignment: 'center' },
                { text: l.headers?.details, bold: true, alignment: 'center' },
                { text: l.headers?.qty, bold: true, alignment: 'center' },
                { text: l.headers?.remarks, bold: true, alignment: 'center' },
              ],
              [
                { text: '1', alignment: 'center', fontSize: 7 },
                { text: '2', alignment: 'center', fontSize: 7 },
                { text: '3', alignment: 'center', fontSize: 7 },
                { text: '4', alignment: 'center', fontSize: 7 },
                { text: '5', alignment: 'center', fontSize: 7 },
              ],
              ...items.map((item, index) => [
                { text: (index + 1).toString(), alignment: 'center' },
                safe(item.goods_description),
                safe(item.goods_details),
                { text: safe(item.quantity) },
                safe(item.remarks),
              ]),
              [
                { text: l.headers?.total, colSpan: 3, alignment: 'right', bold: true },
                {},
                {},
                { text: safe(targetData.tableData?.total_quantity), bold: true, alignment: 'left' },
                {},
              ],
            ],
          },
        },

        {
          margin: [0, 30, 0, 0],
          stack: [
            { text: `${l.footer?.auth_label} : ${safe(targetData.auth_person || '')}`, bold: true },
            { text: `${l.footer?.designation_label} : ${safe(targetData.designation || '')}` },
            { text: `${l.footer?.signature_label} : ____________________`, margin: [0, 5, 0, 0] },
            { text: `${l.footer?.seal_label} :`, margin: [0, 5, 0, 0] },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.4_${lang}.pdf`);
  }

  exportMushak_6_4_Bangla(data: any, lang: string) {
    const l = (data.labels?.mushak_6_4 || {}) as any;
    const targetData = data.mushak_6_4_data?.[lang] || {};
    const items = (targetData.tableData?.rows || []) as any[];

    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');

    const toBnNum = (n: any) => {
      if (!n) return '';
      return n.toString().replace(/\d/g, (d: any) => '০১২৩৪৫৬৭৮৯'[d]);
    };

    (pdfMake as any).fonts = {
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: 'PlaywriteCU', fontSize: 9 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        {
          text: safe(l.titles?.form),
          alignment: 'center',
          bold: true,
          fontSize: 13,
          margin: [0, 5, 0, 0],
        },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 8.5, margin: [0, 0, 0, 15] },

        {
          columns: [
            {
              width: '*',
              stack: [
                {
                  text: `${l.info?.registered_person} : ${safe(targetData.formData?.registered_person)}`,
                },
                {
                  text: `${l.info?.registered_person_bin} : ${safe(targetData.formData?.registered_person_bin)}`,
                },
                {
                  text: `${l.info?.chalan_address} : ${safe(targetData.formData?.chalan_address)}`,
                },
              ],
            },
            {
              width: '25%',
              stack: [
                { text: `${l.info?.inv_no} :`, bold: true },
                { text: `${l.info?.inv_date} :` },
                { text: `${l.info?.inv_time} :` },
              ],
              alignment: 'right',
              margin: [0, 0, 10, 0],
            },
            {
              width: '20%',
              stack: [
                { text: safe(targetData.formData?.chalan_number), bold: true },
                { text: safe(targetData.formData?.issue_date) },
                { text: safe(targetData.formData?.issue_time) },
              ],
              alignment: 'left',
            },
          ],
          margin: [0, 0, 0, 10],
        },

        {
          stack: [
            { text: `${l.info?.buyer_name} : ${safe(targetData.formData?.recipient_name)}` },
            { text: `${l.info?.buyer_bin} : ${safe(targetData.formData?.recipient_bin || '-')}` },
            { text: `${l.info?.dest_addr} : ${safe(targetData.formData?.destination)}` },
          ],
          margin: [0, 0, 0, 15],
        },

        {
          table: {
            headerRows: 2,
            widths: [30, '*', '*', 70, 70],
            body: [
              [
                { text: l.headers?.sl, bold: true, alignment: 'center' },
                { text: l.headers?.desc, bold: true, alignment: 'center' },
                { text: l.headers?.details, bold: true, alignment: 'center' },
                { text: l.headers?.qty, bold: true, alignment: 'center' },
                { text: l.headers?.remarks, bold: true, alignment: 'center' },
              ],
              [
                { text: toBnNum(1), alignment: 'center', fontSize: 8 },
                { text: toBnNum(2), alignment: 'center', fontSize: 8 },
                { text: toBnNum(3), alignment: 'center', fontSize: 8 },
                { text: toBnNum(4), alignment: 'center', fontSize: 8 },
                { text: toBnNum(5), alignment: 'center', fontSize: 8 },
              ],
              ...items.map((item, index) => [
                { text: toBnNum(index + 1), alignment: 'center' },
                safe(item.goods_description),
                safe(item.goods_details),
                { text: safe(item.quantity) },
                safe(item.remarks),
              ]),
              [
                { text: l.headers?.total, colSpan: 3, alignment: 'right', bold: true },
                {},
                {},
                { text: safe(targetData.tableData?.total_quantity), bold: true, alignment: 'left' },
                {},
              ],
            ],
          },
        },

        {
          margin: [0, 30, 0, 0],
          stack: [
            { text: `${l.footer?.auth_label} : ${safe(targetData.auth_person || '')}`, bold: true },
            { text: `${l.footer?.designation_label} : ${safe(targetData.designation || '')}` },
            { text: `${l.footer?.signature_label} : ____________________`, margin: [0, 5, 0, 0] },
            { text: `${l.footer?.seal_label} :`, margin: [0, 5, 0, 0] },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.4_${lang}.pdf`);
  }

  exportMushak_6_5_English(data: any, lang: string) {
    const l = (data.labels?.mushak_6_5 || {}) as any;
    const targetData = data.mushak_6_5_data?.['EN'] || {};
    const items = (targetData.tableData?.rows || []) as any[];

    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: 'Nunito', fontSize: 8.5 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        {
          text: safe(l.titles?.form),
          alignment: 'center',
          bold: true,
          fontSize: 11,
          margin: [0, 5, 0, 0],
        },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 7.5, margin: [0, 0, 0, 15] },

        {
          columns: [
            {
              width: '*',
              stack: [
                {
                  text: `${l.info?.registered_person} : ${safe(targetData.formData?.registered_person)}`,
                },
                {
                  text: `${l.info?.registered_person_bin} : ${safe(targetData.formData?.registered_person_bin)}`,
                },
                { text: `${l.info?.sender_info} : ${safe(targetData.formData?.sender_info)}` },
                { text: `${l.info?.receiver_info} : ${safe(targetData.formData?.receiver_info)}` },
              ],
            },
            {
              width: 'auto',
              stack: [
                {
                  columns: [
                    { text: `${l.info?.inv_no} :`, bold: true, width: 70, alignment: 'right' },
                    { text: safe(targetData.formData?.chalan_number), bold: true, width: '*' },
                  ],
                },
                {
                  columns: [
                    { text: `${l.info?.inv_date} :`, width: 70, alignment: 'right' },
                    { text: safe(targetData.formData?.issue_date), width: '*' },
                  ],
                },
                {
                  columns: [
                    { text: `${l.info?.inv_time} :`, width: 70, alignment: 'right' },
                    { text: safe(targetData.formData?.issue_time), width: '*' },
                  ],
                },
              ],
              margin: [0, 0, 0, 0],
            },
          ],
          margin: [0, 0, 0, 15],
        },

        {
          table: {
            headerRows: 2,
            widths: [25, '*', 50, 60, 60, 60],
            body: [
              [
                { text: l.headers?.sl, bold: true, alignment: 'center' },
                { text: l.headers?.desc, bold: true, alignment: 'center' },
                { text: l.headers?.qty, bold: true, alignment: 'center' },
                { text: l.headers?.price_no_tax, bold: true, alignment: 'center' },
                { text: l.headers?.tax_amount, bold: true, alignment: 'center' },
                { text: l.headers?.remarks, bold: true, alignment: 'center' },
              ],
              [
                { text: '1', alignment: 'center', fontSize: 7 },
                { text: '2', alignment: 'center', fontSize: 7 },
                { text: '3', alignment: 'center', fontSize: 7 },
                { text: '4', alignment: 'center', fontSize: 7 },
                { text: '5', alignment: 'center', fontSize: 7 },
                { text: '6', alignment: 'center', fontSize: 7 },
              ],
              ...items.map((item, index) => [
                { text: (index + 1).toString(), alignment: 'center' },
                safe(item.desc),
                { text: safe(item.qty), alignment: 'center' },
                { text: safe(item.price_no_tax), alignment: 'right' },
                { text: safe(item.tax_amount), alignment: 'right' },
                safe(item.remarks),
              ]),
            ],
          },
        },

        {
          margin: [0, 30, 0, 0],
          stack: [
            { text: `${l.footer?.auth_label}: ____________________`, bold: true },
            { text: `${l.footer?.designation_label}:`, margin: [0, 5, 0, 0] },
            { text: `${l.footer?.signature_label}:`, margin: [0, 5, 0, 0] },
            { text: `${l.footer?.seal_label}:`, margin: [0, 5, 0, 0] },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.5_${lang}.pdf`);
  }

  exportMushak_6_5_Bangla(data: any, lang: string) {
    const l = (data.labels?.mushak_6_5 || {}) as any;
    const targetData = data.mushak_6_5_data?.['BN'] || {};
    const items = (targetData.tableData?.rows || []) as any[];

    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');
    const toBnNum = (n: any) => n.toString().replace(/\d/g, (d: any) => '০১২৩৪৫৬৭৮৯'[d]);

    (pdfMake as any).fonts = {
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: 'PlaywriteCU', fontSize: 9 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        {
          text: safe(l.titles?.form),
          alignment: 'center',
          bold: true,
          fontSize: 12,
          margin: [0, 5, 0, 0],
        },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 8, margin: [0, 0, 0, 15] },

        {
          columns: [
            {
              width: '*',
              stack: [
                {
                  text: `${l.info?.registered_person} : ${safe(targetData.formData?.registered_person)}`,
                },
                {
                  text: `${l.info?.registered_person_bin} : ${safe(targetData.formData?.registered_person_bin)}`,
                },
                { text: `${l.info?.sender_info} : ${safe(targetData.formData?.sender_info)}` },
                { text: `${l.info?.receiver_info} : ${safe(targetData.formData?.receiver_info)}` },
              ],
            },
            {
              width: 'auto',
              stack: [
                {
                  columns: [
                    { text: `${l.info?.inv_no} :`, bold: true, width: 70, alignment: 'right' },
                    { text: safe(targetData.formData?.chalan_number), bold: true, width: '*' },
                  ],
                },
                {
                  columns: [
                    { text: `${l.info?.inv_date} :`, width: 70, alignment: 'right' },
                    { text: safe(targetData.formData?.issue_date), width: '*' },
                  ],
                },
                {
                  columns: [
                    { text: `${l.info?.inv_time} :`, width: 70, alignment: 'right' },
                    { text: safe(targetData.formData?.issue_time), width: '*' },
                  ],
                },
              ],
              margin: [0, 0, 0, 0],
            },
          ],
          margin: [0, 0, 0, 15],
        },

        {
          table: {
            headerRows: 2,
            widths: [30, '*', 50, 60, 60, 50],
            body: [
              [
                { text: l.headers?.sl, bold: true, alignment: 'center' },
                { text: l.headers?.desc, bold: true, alignment: 'center' },
                { text: l.headers?.qty, bold: true, alignment: 'center' },
                { text: l.headers?.price_no_tax, bold: true, alignment: 'center' },
                { text: l.headers?.tax_amount, bold: true, alignment: 'center' },
                { text: l.headers?.remarks, bold: true, alignment: 'center' },
              ],
              [
                { text: toBnNum(1), alignment: 'center', fontSize: 8 },
                { text: toBnNum(2), alignment: 'center', fontSize: 8 },
                { text: toBnNum(3), alignment: 'center', fontSize: 8 },
                { text: toBnNum(4), alignment: 'center', fontSize: 8 },
                { text: toBnNum(5), alignment: 'center', fontSize: 8 },
                { text: toBnNum(6), alignment: 'center', fontSize: 8 },
              ],
              ...items.map((item, index) => [
                { text: toBnNum(index + 1), alignment: 'center' },
                safe(item.desc),
                { text: safe(item.qty), alignment: 'center' },
                { text: safe(item.price_no_tax), alignment: 'right' },
                { text: safe(item.tax_amount), alignment: 'right' },
                safe(item.remarks),
              ]),
            ],
          },
        },

        {
          margin: [0, 30, 0, 0],
          stack: [
            { text: `${l.footer?.auth_label} : ____________________`, bold: true },
            { text: `${l.footer?.designation_label} :`, margin: [0, 5, 0, 0] },
            { text: `${l.footer?.signature_label} :`, margin: [0, 5, 0, 0] },
            { text: `${l.footer?.seal_label} :`, margin: [0, 5, 0, 0] },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.5_${lang}.pdf`);
  }

  exportMushak_6_6_Bangla(data: any, lang: string) {
    const l = (data.labels?.mushak_6_6 || {}) as any;
    const targetData = data.mushak_6_6_data?.BN || {};
    const items = (targetData.tableData?.rows || []) as any[];

    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');
    const toBnNum = (n: any) => n.toString().replace(/\d/g, (d: any) => '০১২৩৪৫৬৭৮৯'[d]);

    (pdfMake as any).fonts = {
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: 'PlaywriteCU', fontSize: 9 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        {
          text: safe(l.titles?.form),
          alignment: 'center',
          bold: true,
          fontSize: 12,
          margin: [0, 5, 0, 0],
        },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 8, margin: [0, 0, 0, 15] },

        {
          columns: [
            {
              width: '*',
              stack: [
                {
                  text: `${l.info?.withholding_entity} : ${safe(targetData.formData?.withholding_entity)}`,
                },
                {
                  text: `${l.info?.withholding_address} : ${safe(targetData.formData?.withholding_address)}`,
                },
                {
                  text: `${l.info?.withholding_bin} : ${safe(targetData.formData?.withholding_bin)}`,
                },
              ],
            },
            {
              width: 'auto',
              stack: [
                { text: `${l.info?.cert_no} : ${safe(targetData.formData?.cert_no)}`, bold: true },
                { text: `${l.info?.issue_date} : ${safe(targetData.formData?.issue_date)}` },
              ],
              alignment: 'right',
            },
          ],
          margin: [0, 0, 0, 15],
        },

        { text: `${l.info?.notes}`, margin: [0, 0, 0, 15] },

        {
          table: {
            headerRows: 2,
            widths: [22, '*', 65, 50, 55, 60, 55, 55],
            body: [
              [
                {
                  text: l.headers?.sl,
                  rowSpan: 2,
                  bold: true,
                  alignment: 'center',
                  margin: [0, 8, 0, 0],
                },
                { text: l.headers?.supplier, colSpan: 2, bold: true, alignment: 'center' },
                {},
                { text: l.headers?.invoice_info, colSpan: 2, bold: true, alignment: 'center' },
                {},
                {
                  text: l.headers?.total_value,
                  rowSpan: 2,
                  bold: true,
                  alignment: 'center',
                  margin: [0, 8, 0, 0],
                },
                {
                  text: l.headers?.vat_amount,
                  rowSpan: 2,
                  bold: true,
                  alignment: 'center',
                  margin: [0, 8, 0, 0],
                },
                {
                  text: l.headers?.vds_amount,
                  rowSpan: 2,
                  bold: true,
                  alignment: 'center',
                  margin: [0, 8, 0, 0],
                },
              ],
              [
                {},
                { text: l.headers?.name, bold: true, alignment: 'center' },
                { text: l.headers?.bin, bold: true, alignment: 'center' },
                { text: l.headers?.number, bold: true, alignment: 'center' },
                { text: l.headers?.issueDate, bold: true, alignment: 'center' },
                {},
                {},
                {},
              ],
              [
                { text: '১', alignment: 'center', fontSize: 7 },
                { text: '২', alignment: 'center', fontSize: 7 },
                { text: '৩', alignment: 'center', fontSize: 7 },
                { text: '৪', alignment: 'center', fontSize: 7 },
                { text: '৫', alignment: 'center', fontSize: 7 },
                { text: '৬', alignment: 'center', fontSize: 7 },
                { text: '৭', alignment: 'center', fontSize: 7 },
                { text: '৮', alignment: 'center', fontSize: 7 },
              ],
              ...items.map((item, index) => [
                { text: toBnNum(index + 1), alignment: 'center' },
                safe(item.name),
                { text: safe(item.bin), alignment: 'center' },
                { text: safe(item.invoice_no), alignment: 'center' },
                { text: safe(item.invoice_date), alignment: 'center' },
                { text: safe(item.total_value), alignment: 'right' },
                { text: safe(item.vat_amount), alignment: 'right' },
                { text: safe(item.vds_amount), alignment: 'right' },
              ]),
              [
                { text: `${l.headers?.total}`, colSpan: 5, alignment: 'right', bold: true },
                {},
                {},
                {},
                {},
                { text: safe(targetData.tableData?.total_supply), bold: true, alignment: 'right' },
                { text: safe(targetData.tableData?.total_vat), bold: true, alignment: 'right' },
                { text: safe(targetData.tableData?.total_vds), bold: true, alignment: 'right' },
              ],
            ],
          },
        },

        {
          margin: [0, 30, 0, 0],
          stack: [
            { text: l.footer?.auth_label, bold: true },
            { text: `${l.footer?.signature_label} : ____________________`, margin: [0, 5, 0, 0] },
            {
              text: `${l.footer?.name_label} : ${safe(targetData.formData?.auth_name || '')}`,
              margin: [0, 5, 0, 0],
            },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.6_${lang}.pdf`);
  }

  exportMushak_6_6_English(data: any, lang: string) {
    const l = (data.labels?.mushak_6_6 || {}) as any;
    const targetData = data.mushak_6_6_data?.EN || {};
    const items = (targetData.tableData?.rows || []) as any[];

    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: 'Nunito', fontSize: 8.5 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        {
          text: safe(l.titles?.form),
          alignment: 'center',
          bold: true,
          fontSize: 11,
          margin: [0, 5, 0, 0],
        },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 7.5, margin: [0, 0, 0, 15] },

        {
          columns: [
            {
              width: '*',
              stack: [
                {
                  text: `${l.info?.withholding_entity}: ${safe(targetData.formData?.withholding_entity)}`,
                },
                {
                  text: `${l.info?.withholding_address}: ${safe(targetData.formData?.withholding_address)}`,
                },
                {
                  text: `${l.info?.withholding_bin}: ${safe(targetData.formData?.withholding_bin)}`,
                },
              ],
            },
            {
              width: 'auto',
              stack: [
                { text: `${l.info?.cert_no}: ${safe(targetData.formData?.cert_no)}`, bold: true },
                { text: `${l.info?.issue_date}: ${safe(targetData.formData?.issue_date)}` },
              ],
              alignment: 'right',
            },
          ],
          margin: [0, 0, 0, 15],
        },

        {
          text: `${l.info?.notes}`,
          margin: [0, 0, 0, 10],
          fontSize: 7.5,
        },

        {
          table: {
            headerRows: 2,
            widths: [20, '*', 65, 50, 55, 60, 55, 55],
            body: [
              [
                {
                  text: l.headers?.sl,
                  rowSpan: 2,
                  bold: true,
                  alignment: 'center',
                  margin: [0, 8, 0, 0],
                },
                { text: l.headers?.supplier, colSpan: 2, bold: true, alignment: 'center' },
                {},
                { text: l.headers?.invoice_info, colSpan: 2, bold: true, alignment: 'center' },
                {},
                {
                  text: l.headers?.total_value,
                  rowSpan: 2,
                  bold: true,
                  alignment: 'center',
                  margin: [0, 8, 0, 0],
                },
                {
                  text: l.headers?.vat_amount,
                  rowSpan: 2,
                  bold: true,
                  alignment: 'center',
                  margin: [0, 8, 0, 0],
                },
                {
                  text: l.headers?.vds_amount,
                  rowSpan: 2,
                  bold: true,
                  alignment: 'center',
                  margin: [0, 8, 0, 0],
                },
              ],
              [
                {},
                { text: l.headers?.name, bold: true, alignment: 'center' },
                { text: l.headers?.bin, bold: true, alignment: 'center' },
                { text: l.headers?.number, bold: true, alignment: 'center' },
                { text: l.headers?.issueDate, bold: true, alignment: 'center' },
                {},
                {},
                {},
              ],
              [
                { text: '1', alignment: 'center', fontSize: 7 },
                { text: '2', alignment: 'center', fontSize: 7 },
                { text: '3', alignment: 'center', fontSize: 7 },
                { text: '4', alignment: 'center', fontSize: 7 },
                { text: '5', alignment: 'center', fontSize: 7 },
                { text: '6', alignment: 'center', fontSize: 7 },
                { text: '7', alignment: 'center', fontSize: 7 },
                { text: '8', alignment: 'center', fontSize: 7 },
              ],
              ...items.map((item, index) => [
                { text: (index + 1).toString(), alignment: 'center' },
                safe(item.name),
                { text: safe(item.bin), alignment: 'center' },
                { text: safe(item.invoice_no), alignment: 'center' },
                { text: safe(item.invoice_date), alignment: 'center' },
                { text: safe(item.total_value), alignment: 'right' },
                { text: safe(item.vat_amount), alignment: 'right' },
                { text: safe(item.vds_amount), alignment: 'right' },
              ]),
              [
                { text: 'Total:', colSpan: 5, alignment: 'right', bold: true },
                {},
                {},
                {},
                {},
                { text: safe(targetData.tableData?.total_supply), bold: true, alignment: 'right' },
                { text: safe(targetData.tableData?.total_vat), bold: true, alignment: 'right' },
                { text: safe(targetData.tableData?.total_vds), bold: true, alignment: 'right' },
              ],
            ],
          },
        },

        {
          margin: [0, 30, 0, 0],
          stack: [
            { text: l.footer?.auth_label, bold: true },
            { text: `${l.footer?.signature_label}: ____________________`, margin: [0, 5, 0, 0] },
            {
              text: `${l.footer?.name_label}: ${safe(targetData.formData?.auth_name || '')}`,
              margin: [0, 5, 0, 0],
            },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.6_${lang}.pdf`);
  }

  exportMushak_6_7_English(data: any, lang: string) {
    const l = (data.labels?.mushak_6_7 || {}) as any;
    const targetData = data.mushak_6_7_data[lang] || {};
    const items = (targetData.tableData?.rows || []) as any[];
    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: 'Nunito', fontSize: 8.5 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        {
          text: safe(l.titles?.form),
          alignment: 'center',
          bold: true,
          fontSize: 11,
          margin: [0, 5, 0, 0],
        },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 7.5, margin: [0, 0, 0, 15] },

        {
          columns: [
            {
              width: '*',
              stack: [
                {
                  text: safe(l.info?.seller_title),
                  bold: true,
                  decoration: 'underline',
                  margin: [0, 0, 0, 5],
                },
                { text: `${safe(l.info?.seller_name)}: ${safe(targetData.formData?.seller_name)}` },
                { text: `${safe(l.info?.seller_bin)}: ${safe(targetData.formData?.seller_bin)}` },
                { text: `${safe(l.info?.orig_inv_no)}: ${safe(targetData.formData?.orig_inv_no)}` },
                {
                  text: `${safe(l.info?.orig_inv_date)}: ${safe(targetData.formData?.orig_inv_date)}`,
                },
              ],
            },
            {
              width: 'auto',
              stack: [
                {
                  text: safe(l.info?.recipient_title),
                  bold: true,
                  decoration: 'underline',
                  margin: [0, 0, 0, 5],
                },
                {
                  text: `${safe(l.info?.recipient_name)}: ${safe(targetData.formData?.recipient_name)}`,
                },
                {
                  text: `${safe(l.info?.recipient_bin)}: ${safe(targetData.formData?.recipient_bin)}`,
                },
                { text: '\n' },
                {
                  text: `${safe(l.info?.credit_note_no)}: ${safe(targetData.formData?.credit_note_no)}`,
                  bold: true,
                },
                { text: `${safe(l.info?.inv_date)}: ${safe(targetData.formData?.issue_date)}` },
                { text: `${safe(l.info?.inv_time)}: ${safe(targetData.formData?.issue_time)}` },
              ],
              alignment: 'left',
            },
          ],
          margin: [0, 0, 0, 15],
        },

        {
          table: {
            headerRows: 1,
            widths: [25, '*', 60, 60, 70, 80],
            body: [
              [
                { text: l.headers?.sl, bold: true, alignment: 'center' },
                { text: l.headers?.details, bold: true, alignment: 'center' },
                { text: l.headers?.unit, bold: true, alignment: 'center' },
                { text: l.headers?.qty, bold: true, alignment: 'center' },
                { text: l.headers?.u_price, bold: true, alignment: 'center' },
                { text: l.headers?.t_price, bold: true, alignment: 'center' },
              ],
              ...items.map((item, index) => [
                { text: (index + 1).toString(), alignment: 'center' },
                safe(item.details),
                { text: safe(item.unit), alignment: 'center' },
                { text: safe(item.qty), alignment: 'center' },
                { text: safe(item.u_price), alignment: 'right' },
                { text: safe(item.t_price), alignment: 'right' },
              ]),
              [
                { text: safe(l.summary?.total_val), colSpan: 5, alignment: 'right', bold: true },
                {},
                {},
                {},
                {},
                { text: safe(targetData.tableData?.total_val), alignment: 'right', bold: true },
              ],
            ],
          },
        },
        {
          columns: [
            { width: '*', text: '' },
            {
              width: 'auto',
              table: {
                widths: [160, 80],
                body: [
                  [
                    { text: l.summary?.deduction, alignment: 'left' },
                    { text: safe(targetData.tableData?.deduction), alignment: 'right' },
                  ],
                  [
                    { text: l.summary?.val_incl_vat, alignment: 'left' },
                    { text: safe(targetData.tableData?.val_incl_vat), alignment: 'right' },
                  ],
                  [
                    { text: l.summary?.vat_amount, alignment: 'left' },
                    { text: safe(targetData.tableData?.vat_amount), alignment: 'right' },
                  ],
                  [
                    { text: l.summary?.sd_amount, alignment: 'left' },
                    { text: safe(targetData.tableData?.sd_amount), alignment: 'right' },
                  ],
                  [
                    { text: l.summary?.total_tax, alignment: 'left', bold: true },
                    { text: safe(targetData.tableData?.total_tax), alignment: 'right', bold: true },
                  ],
                ],
              },
              margin: [0, 0, 0, 10],
            },
          ],
        },

        // Reasons Box
        { text: l.headers?.reasons, bold: true, margin: [0, 5, 0, 2] },
        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: safe(targetData.tableData?.reasons),
                  minHeight: 150,
                  margin: [5, 5, 5, 5],
                },
              ],
            ],
          },
          margin: [0, 0, 0, 40],
        },

        // Signature
        { text: safe(l.footer?.auth_label), alignment: 'right', bold: true, margin: [0, 0, 0, 30] },

        // Notes
        { text: safe(l.footer?.unitPrice), fontSize: 7.5 },
        { text: safe(l.footer?.deduction), fontSize: 7.5 },
        { text: safe(l.footer?.totalTax), fontSize: 7.5 },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.7_${lang}.pdf`);
  }

  exportMushak_6_7_Bangla(data: any, lang: string) {
    const l = (data.labels?.mushak_6_7 || {}) as any;
    const targetData = data.mushak_6_7_data[lang] || {};
    const items = (targetData.tableData?.rows || []) as any[];
    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');

    (pdfMake as any).fonts = {
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: 'PlaywriteCU', fontSize: 8.5 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        {
          text: safe(l.titles?.form),
          alignment: 'center',
          bold: true,
          fontSize: 11,
          margin: [0, 5, 0, 0],
        },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 7.5, margin: [0, 0, 0, 15] },

        {
          columns: [
            {
              width: '*',
              stack: [
                {
                  text: safe(l.info?.seller_title),
                  bold: true,
                  decoration: 'underline',
                  margin: [0, 0, 0, 5],
                },
                { text: `${safe(l.info?.seller_name)}: ${safe(targetData.formData?.seller_name)}` },
                { text: `${safe(l.info?.seller_bin)}: ${safe(targetData.formData?.seller_bin)}` },
                { text: `${safe(l.info?.orig_inv_no)}: ${safe(targetData.formData?.orig_inv_no)}` },
                {
                  text: `${safe(l.info?.orig_inv_date)}: ${safe(targetData.formData?.orig_inv_date)}`,
                },
              ],
            },
            {
              width: 'auto',
              stack: [
                {
                  text: safe(l.info?.recipient_title),
                  bold: true,
                  decoration: 'underline',
                  margin: [0, 0, 0, 5],
                },
                {
                  text: `${safe(l.info?.recipient_name)}: ${safe(targetData.formData?.recipient_name)}`,
                },
                {
                  text: `${safe(l.info?.recipient_bin)}: ${safe(targetData.formData?.recipient_bin)}`,
                },
                { text: '\n' },
                {
                  text: `${safe(l.info?.credit_note_no)}: ${safe(targetData.formData?.credit_note_no)}`,
                  bold: true,
                },
                { text: `${safe(l.info?.inv_date)}: ${safe(targetData.formData?.issue_date)}` },
                { text: `${safe(l.info?.inv_time)}: ${safe(targetData.formData?.issue_time)}` },
              ],
              alignment: 'left',
            },
          ],
          margin: [0, 0, 0, 15],
        },

        {
          table: {
            headerRows: 1,
            widths: [25, '*', 60, 60, 70, 80],
            body: [
              [
                { text: l.headers?.sl, bold: true, alignment: 'center' },
                { text: l.headers?.details, bold: true, alignment: 'center' },
                { text: l.headers?.unit, bold: true, alignment: 'center' },
                { text: l.headers?.qty, bold: true, alignment: 'center' },
                { text: l.headers?.u_price, bold: true, alignment: 'center' },
                { text: l.headers?.t_price, bold: true, alignment: 'center' },
              ],
              ...items.map((item, index) => [
                { text: (index + 1).toString(), alignment: 'center' },
                safe(item.details),
                { text: safe(item.unit), alignment: 'center' },
                { text: safe(item.qty), alignment: 'center' },
                { text: safe(item.u_price), alignment: 'right' },
                { text: safe(item.t_price), alignment: 'right' },
              ]),
              [
                { text: safe(l.summary?.total_val), colSpan: 5, alignment: 'right', bold: true },
                {},
                {},
                {},
                {},
                { text: safe(targetData.tableData?.total_val), alignment: 'right', bold: true },
              ],
            ],
          },
        },
        {
          columns: [
            { width: '*', text: '' },
            {
              width: 'auto',
              table: {
                widths: [160, 80],
                body: [
                  [
                    { text: l.summary?.deduction, alignment: 'left' },
                    { text: safe(targetData.tableData?.deduction), alignment: 'right' },
                  ],
                  [
                    { text: l.summary?.val_incl_vat, alignment: 'left' },
                    { text: safe(targetData.tableData?.val_incl_vat), alignment: 'right' },
                  ],
                  [
                    { text: l.summary?.vat_amount, alignment: 'left' },
                    { text: safe(targetData.tableData?.vat_amount), alignment: 'right' },
                  ],
                  [
                    { text: l.summary?.sd_amount, alignment: 'left' },
                    { text: safe(targetData.tableData?.sd_amount), alignment: 'right' },
                  ],
                  [
                    { text: l.summary?.total_tax, alignment: 'left', bold: true },
                    { text: safe(targetData.tableData?.total_tax), alignment: 'right', bold: true },
                  ],
                ],
              },
              margin: [0, 0, 0, 10],
            },
          ],
        },

        // Reasons Box
        { text: l.headers?.reasons, bold: true, margin: [0, 5, 0, 2] },
        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: safe(targetData.tableData?.reasons),
                  minHeight: 150,
                  margin: [5, 5, 5, 5],
                },
              ],
            ],
          },
          margin: [0, 0, 0, 40],
        },

        // Signature
        { text: safe(l.footer?.auth_label), alignment: 'right', bold: true, margin: [0, 0, 0, 30] },

        // Notes
        { text: safe(l.footer?.unitPrice), fontSize: 7.5 },
        { text: safe(l.footer?.deduction), fontSize: 7.5 },
        { text: safe(l.footer?.totalTax), fontSize: 7.5 },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.7_${lang}.pdf`);
  }

  exportMushak_6_8_English(data: any, lang: string) {
    const l = (data.labels?.mushak_6_8 || {}) as any;
    const targetData = data.mushak_6_8_data[lang] || {};
    const items = (targetData.tableData?.rows || []) as any[];
    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: 'Nunito', fontSize: 8.5 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        {
          text: safe(l.titles?.form),
          alignment: 'center',
          bold: true,
          fontSize: 11,
          margin: [0, 5, 0, 0],
        },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 7.5, margin: [0, 0, 0, 15] },

        {
          columns: [
            {
              width: '*',
              stack: [
                {
                  text: safe(l.info?.seller_title),
                  bold: true,
                  decoration: 'underline',
                  margin: [0, 0, 0, 5],
                },
                { text: `${safe(l.info?.name)}: ${safe(targetData.formData?.seller_name)}` },
                { text: `${safe(l.info?.bin)}: ${safe(targetData.formData?.seller_bin)}` },
                { text: `${safe(l.info?.orig_inv_no)}: ${safe(targetData.formData?.orig_inv_no)}` },
                {
                  text: `${safe(l.info?.orig_inv_date)}: ${safe(targetData.formData?.orig_inv_date)}`,
                },
              ],
            },
            {
              width: 'auto',
              stack: [
                {
                  text: safe(l.info?.recipient_title),
                  bold: true,
                  decoration: 'underline',
                  margin: [0, 0, 0, 5],
                },
                {
                  text: `${safe(l.info?.recipient_name)}: ${safe(targetData.formData?.recipient_name)}`,
                },
                {
                  text: `${safe(l.info?.recipient_bin)}: ${safe(targetData.formData?.recipient_bin)}`,
                },
                { text: '\n' },
                {
                  text: `${safe(l.info?.credit_note_no)}: ${safe(targetData.formData?.credit_note_no)}`,
                  bold: true,
                },
                { text: `${safe(l.info?.inv_date)}: ${safe(targetData.formData?.issue_date)}` },
                { text: `${safe(l.info?.inv_time)}: ${safe(targetData.formData?.issue_time)}` },
              ],
              alignment: 'left',
            },
          ],
          margin: [0, 0, 0, 15],
        },

        {
          table: {
            headerRows: 1,
            widths: [25, '*', 60, 60, 70, 80],
            body: [
              [
                { text: l.headers?.sl, bold: true, alignment: 'center' },
                { text: l.headers?.details, bold: true, alignment: 'center' },
                { text: l.headers?.unit, bold: true, alignment: 'center' },
                { text: l.headers?.qty, bold: true, alignment: 'center' },
                { text: l.headers?.u_price, bold: true, alignment: 'center' },
                { text: l.headers?.t_price, bold: true, alignment: 'center' },
              ],
              ...items.map((item, index) => [
                { text: (index + 1).toString(), alignment: 'center' },
                safe(item.details),
                { text: safe(item.unit), alignment: 'center' },
                { text: safe(item.qty), alignment: 'center' },
                { text: safe(item.u_price), alignment: 'right' },
                { text: safe(item.t_price), alignment: 'right' },
              ]),
              [
                { text: safe(l.summary?.total_val), colSpan: 5, alignment: 'right', bold: true },
                {},
                {},
                {},
                {},
                { text: safe(targetData.tableData?.total_val), alignment: 'right', bold: true },
              ],
            ],
          },
        },
        {
          columns: [
            { width: '*', text: '' },
            {
              width: 'auto',
              table: {
                widths: [160, 80],
                body: [
                  [
                    { text: l.summary?.deduction, alignment: 'left' },
                    { text: safe(targetData.tableData?.deduction), alignment: 'right' },
                  ],
                  [
                    { text: l.summary?.val_incl_vat, alignment: 'left' },
                    { text: safe(targetData.tableData?.val_incl_vat), alignment: 'right' },
                  ],
                  [
                    { text: l.summary?.vat_amount, alignment: 'left' },
                    { text: safe(targetData.tableData?.vat_amount), alignment: 'right' },
                  ],
                  [
                    { text: l.summary?.sd_amount, alignment: 'left' },
                    { text: safe(targetData.tableData?.sd_amount), alignment: 'right' },
                  ],
                  [
                    { text: l.summary?.total_tax, alignment: 'left', bold: true },
                    { text: safe(targetData.tableData?.total_tax), alignment: 'right', bold: true },
                  ],
                ],
              },
              margin: [0, 0, 0, 10],
            },
          ],
        },

        // Reasons Box
        { text: l.headers?.reasons, bold: true, margin: [0, 5, 0, 2] },
        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: safe(targetData.tableData?.reasons),
                  minHeight: 150,
                  margin: [5, 5, 5, 5],
                },
              ],
            ],
          },
          margin: [0, 0, 0, 40],
        },

        // Signature
        { text: safe(l.footer?.auth_label), alignment: 'right', bold: true, margin: [0, 0, 0, 30] },

        // Notes
        { text: safe(l.footer?.unitPrice), fontSize: 7.5 },
        { text: safe(l.footer?.deduction), fontSize: 7.5 },
        { text: safe(l.footer?.totalTax), fontSize: 7.5 },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.8_${lang}.pdf`);
  }

  exportMushak_6_8_Bangla(data: any, lang: string) {
    const l = (data.labels?.mushak_6_8 || {}) as any;
    const targetData = data.mushak_6_8_data[lang] || {};
    const items = (targetData.tableData?.rows || []) as any[];
    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');

    (pdfMake as any).fonts = {
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: 'PlaywriteCU', fontSize: 8.5 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        {
          text: safe(l.titles?.form),
          alignment: 'center',
          bold: true,
          fontSize: 11,
          margin: [0, 5, 0, 0],
        },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 7.5, margin: [0, 0, 0, 15] },

        {
          columns: [
            {
              width: '*',
              stack: [
                {
                  text: safe(l.info?.seller_title),
                  bold: true,
                  decoration: 'underline',
                  margin: [0, 0, 0, 5],
                },
                { text: `${safe(l.info?.name)}: ${safe(targetData.formData?.seller_name)}` },
                { text: `${safe(l.info?.bin)}: ${safe(targetData.formData?.seller_bin)}` },
                { text: `${safe(l.info?.orig_inv_no)}: ${safe(targetData.formData?.orig_inv_no)}` },
                {
                  text: `${safe(l.info?.orig_inv_date)}: ${safe(targetData.formData?.orig_inv_date)}`,
                },
              ],
            },
            {
              width: 'auto',
              stack: [
                {
                  text: safe(l.info?.recipient_title),
                  bold: true,
                  decoration: 'underline',
                  margin: [0, 0, 0, 5],
                },
                {
                  text: `${safe(l.info?.recipient_name)}: ${safe(targetData.formData?.recipient_name)}`,
                },
                {
                  text: `${safe(l.info?.recipient_bin)}: ${safe(targetData.formData?.recipient_bin)}`,
                },
                { text: '\n' },
                {
                  text: `${safe(l.info?.credit_note_no)}: ${safe(targetData.formData?.credit_note_no)}`,
                  bold: true,
                },
                { text: `${safe(l.info?.inv_date)}: ${safe(targetData.formData?.issue_date)}` },
                { text: `${safe(l.info?.inv_time)}: ${safe(targetData.formData?.issue_time)}` },
              ],
              alignment: 'left',
            },
          ],
          margin: [0, 0, 0, 15],
        },

        {
          table: {
            headerRows: 1,
            widths: [25, '*', 60, 60, 70, 80],
            body: [
              [
                { text: l.headers?.sl, bold: true, alignment: 'center' },
                { text: l.headers?.details, bold: true, alignment: 'center' },
                { text: l.headers?.unit, bold: true, alignment: 'center' },
                { text: l.headers?.qty, bold: true, alignment: 'center' },
                { text: l.headers?.u_price, bold: true, alignment: 'center' },
                { text: l.headers?.t_price, bold: true, alignment: 'center' },
              ],
              ...items.map((item, index) => [
                { text: (index + 1).toString(), alignment: 'center' },
                safe(item.details),
                { text: safe(item.unit), alignment: 'center' },
                { text: safe(item.qty), alignment: 'center' },
                { text: safe(item.u_price), alignment: 'right' },
                { text: safe(item.t_price), alignment: 'right' },
              ]),
              [
                { text: safe(l.summary?.total_val), colSpan: 5, alignment: 'right', bold: true },
                {},
                {},
                {},
                {},
                { text: safe(targetData.tableData?.total_val), alignment: 'right', bold: true },
              ],
            ],
          },
        },
        {
          columns: [
            { width: '*', text: '' },
            {
              width: 'auto',
              table: {
                widths: [160, 80],
                body: [
                  [
                    { text: l.summary?.deduction, alignment: 'left' },
                    { text: safe(targetData.tableData?.deduction), alignment: 'right' },
                  ],
                  [
                    { text: l.summary?.val_incl_vat, alignment: 'left' },
                    { text: safe(targetData.tableData?.val_incl_vat), alignment: 'right' },
                  ],
                  [
                    { text: l.summary?.vat_amount, alignment: 'left' },
                    { text: safe(targetData.tableData?.vat_amount), alignment: 'right' },
                  ],
                  [
                    { text: l.summary?.sd_amount, alignment: 'left' },
                    { text: safe(targetData.tableData?.sd_amount), alignment: 'right' },
                  ],
                  [
                    { text: l.summary?.total_tax, alignment: 'left', bold: true },
                    { text: safe(targetData.tableData?.total_tax), alignment: 'right', bold: true },
                  ],
                ],
              },
              margin: [0, 0, 0, 10],
            },
          ],
        },

        // Reasons Box
        { text: l.headers?.reasons, bold: true, margin: [0, 5, 0, 2] },
        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: safe(targetData.tableData?.reasons),
                  minHeight: 150,
                  margin: [5, 5, 5, 5],
                },
              ],
            ],
          },
          margin: [0, 0, 0, 40],
        },

        // Signature
        { text: safe(l.footer?.auth_label), alignment: 'right', bold: true, margin: [0, 0, 0, 30] },

        // Notes
        { text: safe(l.footer?.unitPrice), fontSize: 7.5 },
        { text: safe(l.footer?.deduction), fontSize: 7.5 },
        { text: safe(l.footer?.totalTax), fontSize: 7.5 },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.8_${lang}.pdf`);
  }

  exportMushak_6_9_English(data: any, lang: string) {
    const l = (data.labels?.mushak_6_9 || {}) as any;
    const targetData = data.mushak_6_9_data[lang] || {};
    const items = (targetData.tableData?.rows || []) as any[];
    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: 'Nunito', fontSize: 9 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        {
          text: safe(l.titles?.form),
          alignment: 'center',
          bold: true,
          fontSize: 11,
          margin: [0, 5, 0, 0],
        },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 8, margin: [0, 0, 0, 15] },

        {
          columns: [
            {
              width: '*',
              stack: [
                {
                  text: `${safe(l.info?.listed_person)}: ${safe(targetData.formData?.listed_person)}`,
                },
                { text: `${safe(l.info?.listed_bin)}: ${safe(targetData.formData?.listed_bin)}` },
                {
                  text: `${safe(l.info?.issue_address)}: ${safe(targetData.formData?.issue_address)}`,
                },
                {
                  text: `${safe(l.info?.inv_no)}: ${safe(targetData.formData?.chalan_number)}`,
                  bold: true,
                },
                { text: `${safe(l.info?.inv_date)}: ${safe(targetData.formData?.issue_date)}` },
              ],
            },
            {
              width: 'auto',
              stack: [
                { text: `${safe(l.info?.buyer_name)}: ${safe(targetData.formData?.buyer_name)}` },
                { text: `${safe(l.info?.buyer_bin)}: ${safe(targetData.formData?.buyer_bin)}` },
              ],
              alignment: 'right',
            },
          ],
          margin: [0, 0, 0, 15],
        },

        {
          table: {
            headerRows: 1,
            widths: [35, '*', 60, 60, 80, 80],
            body: [
              [
                { text: l.headers?.sl, bold: true, alignment: 'center' },
                { text: l.headers?.desc, bold: true, alignment: 'center' },
                { text: l.headers?.unit, bold: true, alignment: 'center' },
                { text: l.headers?.qty, bold: true, alignment: 'center' },
                { text: l.headers?.u_price, bold: true, alignment: 'center' },
                { text: l.headers?.t_price, bold: true, alignment: 'center' },
              ],
              ...items.map((item, index) => [
                { text: (index + 1).toString(), alignment: 'center' },
                safe(item.desc),
                { text: safe(item.unit), alignment: 'center' },
                { text: safe(item.qty), alignment: 'center' },
                { text: safe(item.u_price), alignment: 'right' },
                { text: safe(item.t_price), alignment: 'right' },
              ]),
              [
                { text: safe(l.headers?.grand_total), colSpan: 5, alignment: 'right', bold: true },
                {},
                {},
                {},
                {},
                { text: safe(targetData.tableData?.grand_total), alignment: 'right', bold: true },
              ],
              [
                { text: safe(l.headers?.turnover_tax), colSpan: 5, alignment: 'right', bold: true },
                {},
                {},
                {},
                {},
                { text: safe(targetData.tableData?.turnover_tax), alignment: 'right', bold: true },
              ],
            ],
          },
        },

        {
          margin: [0, 25, 0, 0],
          stack: [
            { text: safe(l.notes?.note1), fontSize: 8 },
            { text: safe(l.notes?.note2), fontSize: 8, margin: [0, 5, 0, 0] },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.9_${lang}.pdf`);
  }

  exportMushak_6_9_Bangla(data: any, lang: string) {
    const l = (data.labels?.mushak_6_9 || {}) as any;
    const targetData = data.mushak_6_9_data[lang] || {};
    const items = (targetData.tableData?.rows || []) as any[];
    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');
    const toBnNum = (n: any) =>
      lang === 'BN' ? n.toString().replace(/\d/g, (d: any) => '০১২৩৪৫৬৭৮৯'[d]) : n;

    (pdfMake as any).fonts = {
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: 'PlaywriteCU', fontSize: 9 },
      content: [
        { text: safe(l.titles?.m_name), alignment: 'right', bold: true },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        { text: safe(l.titles?.form), alignment: 'center', bold: true, fontSize: 12 },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 8.5, margin: [0, 2, 0, 15] },

        {
          columns: [
            {
              width: '*',
              stack: [
                {
                  text: `${safe(l.info?.listed_person)} : ${safe(targetData.formData?.listed_person)}`,
                },
                { text: `${safe(l.info?.listed_bin)} : ${safe(targetData.formData?.listed_bin)}` },
                {
                  text: `${safe(l.info?.issue_address)} : ${safe(targetData.formData?.issue_address)}`,
                },
                {
                  text: `${safe(l.info?.inv_no)} : ${safe(targetData.formData?.chalan_number)}`,
                  bold: true,
                },
                { text: `${safe(l.info?.inv_date)} : ${safe(targetData.formData?.issue_date)}` },
              ],
            },
            {
              width: 'auto',
              stack: [
                { text: `${safe(l.info?.buyer_name)} : ${safe(targetData.formData?.buyer_name)}` },
                { text: `${safe(l.info?.buyer_bin)} : ${safe(targetData.formData?.buyer_bin)}` },
              ],
              alignment: 'right',
            },
          ],
          margin: [0, 0, 0, 15],
        },

        {
          table: {
            headerRows: 1,
            widths: [35, '*', 60, 60, 80, 80],
            body: [
              [
                { text: l.headers?.sl, bold: true, alignment: 'center' },
                { text: l.headers?.desc, bold: true, alignment: 'center' },
                { text: l.headers?.unit, bold: true, alignment: 'center' },
                { text: l.headers?.qty, bold: true, alignment: 'center' },
                { text: l.headers?.u_price, bold: true, alignment: 'center' },
                { text: l.headers?.t_price, bold: true, alignment: 'center' },
              ],
              ...items.map((item, index) => [
                { text: toBnNum(index + 1), alignment: 'center' },
                safe(item.desc),
                { text: safe(item.unit), alignment: 'center' },
                { text: toBnNum(safe(item.qty)), alignment: 'center' },
                { text: toBnNum(safe(item.u_price)), alignment: 'right' },
                { text: toBnNum(safe(item.t_price)), alignment: 'right' },
              ]),
              [
                { text: safe(l.headers?.grand_total), colSpan: 5, alignment: 'right', bold: true },
                {},
                {},
                {},
                {},
                {
                  text: toBnNum(safe(targetData.tableData?.grand_total)),
                  alignment: 'right',
                  bold: true,
                },
              ],
              [
                { text: safe(l.headers?.turnover_tax), colSpan: 5, alignment: 'right', bold: true },
                {},
                {},
                {},
                {},
                {
                  text: toBnNum(safe(targetData.tableData?.turnover_tax)),
                  alignment: 'right',
                  bold: true,
                },
              ],
            ],
          },
        },

        {
          margin: [0, 25, 0, 0],
          stack: [
            { text: safe(l.notes?.note1), fontSize: 8 },
            { text: safe(l.notes?.note2), fontSize: 8, margin: [0, 5, 0, 0] },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.9_${lang}.pdf`);
  }

  exportMushak_6_10_Bangla(data: any, lang: string) {
    const l = (data.labels?.mushak_6_10 || {}) as any;
    const targetData = data.mushak_6_10_data[lang] || {};
    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');

    const toBnNum = (n: any) =>
      lang === 'bl' ? n.toString().replace(/\d/g, (d: any) => '০১২৩৪৫৬৭৮৯'[d]) : n;

    (pdfMake as any).fonts = {
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: 'PlaywriteCU', fontSize: 8 },
      content: [
        { text: safe(l.titles?.gov), alignment: 'center' },
        { text: safe(l.titles?.nbr), alignment: 'center' },
        {
          text: safe(l.titles?.sub),
          alignment: 'center',
          bold: true,
          fontSize: 9,
          margin: [0, 5, 0, 0],
        },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 7, margin: [0, 2, 0, 10] },
        { text: safe(l.titles?.form), alignment: 'right', bold: true },

        { text: `${l.info?.name} ${safe(targetData.formData?.person_name)}` },
        {
          text: `${l.info?.bin} ${toBnNum(safe(targetData.formData?.bin))}`,
          margin: [0, 0, 0, 15],
        },

        // Part-A: Purchase [cite: 46-47]
        { text: l.sections?.part_a, bold: true, margin: [0, 5, 0, 5], decoration: 'underline' },
        {
          table: {
            headerRows: 1,
            widths: [20, 50, 55, 60, '*', '*', 70],
            body: [
              [
                l.headers?.sl,
                l.headers?.inv_no,
                l.headers?.inv_date,
                l.headers?.value,
                l.headers?.seller_name,
                l.headers?.seller_address,
                l.headers?.seller_id,
              ].map((h) => ({ text: h, alignment: 'center' })),
              ...(targetData.purchaseData || []).map((p: any, i: number) => [
                { text: toBnNum(i + 1), alignment: 'center' }, // ১, ২, ৩ নিশ্চিত করবে
                safe(p.inv_no),
                toBnNum(safe(p.date)),
                { text: toBnNum(safe(p.val)), alignment: 'right' },
                safe(p.name),
                safe(p.addr),
                toBnNum(safe(p.id)),
              ]),
            ],
          },
        },

        // Part-B: Sales [cite: 48-49]
        { text: l.sections?.part_b, bold: true, margin: [0, 15, 0, 5], decoration: 'underline' },
        {
          table: {
            headerRows: 1,
            widths: [20, 50, 55, 60, '*', '*', 70],
            body: [
              [
                l.headers?.sl,
                l.headers?.inv_no,
                l.headers?.inv_date,
                l.headers?.value,
                l.headers?.buyer_name,
                l.headers?.buyer_address,
                l.headers?.buyer_id,
              ].map((h) => ({ text: h, alignment: 'center' })),
              ...(targetData.salesData || []).map((s: any, i: number) => [
                { text: toBnNum(i + 1), alignment: 'center' },
                safe(s.inv_no),
                toBnNum(safe(s.date)),
                { text: toBnNum(safe(s.val)), alignment: 'right' },
                safe(s.name),
                safe(s.addr),
                toBnNum(safe(s.id)),
              ]),
            ],
          },
        },

        {
          text: '\nদায়িত্বপ্রাপ্ত ব্যক্তির স্বাক্ষর: ____________________',
          margin: [0, 30, 0, 0],
        },
        { text: `নাম: ${safe(targetData.formData?.person_name)}` },
        { text: `তারিখঃ ${toBnNum(new Date().toLocaleDateString())}` },
        {
          text: safe(l.notes?.special_note),
          fontSize: 7.5,
          italics: true,
          margin: [0, 20, 0, 0],
          alignment: 'justify',
          bold: true,
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.10_${lang}.pdf`);
  }

  exportMushak_6_10_English(data: any, lang: string) {
    debugger;
    const l = (data.labels?.mushak_6_10 || {}) as any;
    const targetData = data.mushak_6_10_data[lang] || {};
    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: 'Nunito', fontSize: 8.5 },
      content: [
        { text: safe(l.titles?.gov), alignment: 'center' },
        { text: safe(l.titles?.nbr), alignment: 'center' },
        {
          text: safe(l.titles?.sub),
          alignment: 'center',
          bold: true,
          fontSize: 9.5,
          margin: [0, 5, 0, 0],
        },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 7, margin: [0, 2, 0, 10] },
        { text: safe(l.titles?.form), alignment: 'right', bold: true, fontSize: 10 },

        { text: `${l.info?.name} ${safe(targetData.formData?.person_name)}` },
        { text: `${l.info?.bin} ${safe(targetData.formData?.bin)}`, margin: [0, 0, 0, 15] },

        // Part-A: Purchase Info [cite: 46-47]
        {
          text: safe(l.sections?.part_a),
          bold: true,
          margin: [0, 5, 0, 5],
          decoration: 'underline',
        },
        {
          table: {
            headerRows: 1,
            widths: [25, 55, 60, 65, '*', '*', 75],
            body: [
              [
                { text: l.headers?.sl, alignment: 'center' },
                { text: l.headers?.inv_no, alignment: 'center' },
                { text: l.headers?.inv_date, alignment: 'center' },
                { text: l.headers?.value, alignment: 'center' },
                { text: l.headers?.seller_name, alignment: 'center' },
                { text: l.headers?.seller_address, alignment: 'center' },
                { text: l.headers?.seller_id, alignment: 'center' },
              ],
              ...(targetData.purchaseData || []).map((p: any, i: number) => [
                { text: (i + 1).toString(), alignment: 'center' },
                safe(p.inv_no),
                safe(p.date),
                { text: safe(p.val), alignment: 'right' },
                safe(p.name),
                safe(p.addr),
                safe(p.id),
              ]),
            ],
          },
        },

        // Part-B: Sales Info [cite: 48-49]
        {
          text: safe(l.sections?.part_b),
          bold: true,
          margin: [0, 15, 0, 5],
          decoration: 'underline',
        },
        {
          table: {
            headerRows: 1,
            widths: [25, 55, 60, 65, '*', '*', 75],
            body: [
              [
                { text: l.headers?.sl, alignment: 'center' },
                { text: l.headers?.inv_no, alignment: 'center' },
                { text: l.headers?.inv_date, alignment: 'center' },
                { text: l.headers?.value, alignment: 'center' },
                { text: l.headers?.buyer_name, alignment: 'center' },
                { text: l.headers?.buyer_address, alignment: 'center' },
                { text: l.headers?.buyer_id, alignment: 'center' },
              ],
              ...(targetData.salesData || []).map((s: any, i: number) => [
                { text: (i + 1).toString(), alignment: 'center' },
                safe(s.inv_no),
                safe(s.date),
                { text: safe(s.val), alignment: 'right' },
                safe(s.name),
                safe(s.addr),
                safe(s.id),
              ]),
            ],
          },
        },

        // Footer Section [cite: 50-52]
        {
          margin: [0, 30, 0, 0],
          stack: [
            { text: `Signature of Officer-in-charge: ____________________`, margin: [0, 0, 0, 5] },
            {
              text: `Name: ${safe(targetData.formData?.auth_name || targetData.formData?.person_name)}`,
            },
            { text: `Date: ${safe(targetData.formData?.date || new Date().toLocaleDateString())}` },
          ],
        },
        {
          text: safe(l.notes?.special_note),
          fontSize: 7.5,
          italics: true,
          margin: [0, 20, 0, 0],
          alignment: 'justify',
          bold: true,
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`Mushak_6.10_${lang}.pdf`);
  }

  exportMushak_10_1(data: any, lang: string) {
    debugger;
    const l = (data.labels?.mushak_10_1 || {}) as any;
    const targetData = data.mushak_10_1_data?.[lang] || {};
    const items = (targetData.tableData || []) as any[];
    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');
    const toBnNum = (n: any) =>
      lang === 'BN' ? n.toString().replace(/\d/g, (d: any) => '০১২৩৪৫৬৭৮৯'[d]) : n;

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: lang === 'BN' ? 'PlaywriteCU' : 'Nunito', fontSize: 8.5 },
      content: [
        {
          columns: [
            { width: '*', text: '' },
            {
              width: 'auto',
              table: {
                widths: [80],
                body: [[{ text: safe(l.titles?.form), alignment: 'center', bold: true }]],
              },
            },
          ],
          // margin: [0, 10, 0, 10]
        },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        {
          text: safe(l.titles?.name),
          alignment: 'center',
          bold: true,
          fontSize: 10,
          margin: [0, 5, 0, 0],
        },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 7.5, margin: [0, 2, 0, 15] },

        // Part-1
        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: safe(l.sections?.part1),
                  bold: true,
                  fillColor: '#eeeeee',
                  alignment: 'center',
                },
              ],
            ],
          },
        },
        {
          table: {
            widths: [200, 10, '*'],
            body: [
              [{ text: l.info?.org_name }, ':', safe(targetData.formData?.org_name)],
              [{ text: l.info?.bin }, ':', toBnNum(safe(targetData.formData?.bin))],
              [{ text: l.info?.applicant_name }, ':', safe(targetData.formData?.applicant_name)],
              [{ text: l.info?.designation }, ':', safe(targetData.formData?.designation)],
              [{ text: l.info?.id_no }, ':', toBnNum(safe(targetData.formData?.id_no))],
            ],
          },
          margin: [0, 0, 0, 15],
        },

        // Part-2
        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: safe(l.sections?.part2),
                  bold: true,
                  fillColor: '#eeeeee',
                  alignment: 'center',
                },
              ],
            ],
          },
        },
        {
          table: {
            widths: [200, 10, '*'],
            body: [
              [
                { text: l.info?.actual_vat, bold: true },
                ':',
                { text: toBnNum(safe(targetData.total_vat)), bold: true },
              ],
              [
                { text: l.info?.attachments },
                ':',
                {
                  table: {
                    widths: [15, '*', 60],
                    body: [
                      [
                        {
                          text: l.info?.copy_no,
                          colSpan: 3,
                          alignment: 'right',
                          fontSize: 7,
                          margin: [0, 0, 5, 0],
                          border: [false, false, false, false],
                        },
                        {},
                        {},
                      ],
                      [
                        { text: '(a)', alignment: 'center' },
                        l.info?.att_a,
                        { text: targetData.formData?.att_a, alignment: 'center' },
                      ],
                      [
                        { text: '(b)', alignment: 'center' },
                        l.info?.att_b,
                        { text: targetData.formData?.att_b, alignment: 'center' },
                      ],
                      [
                        { text: '(c)', alignment: 'center' },
                        l.info?.att_c,
                        { text: targetData.formData?.att_c, alignment: 'center' },
                      ],
                      [
                        { text: '(d)', alignment: 'center' },
                        l.info?.att_d,
                        { text: targetData.formData?.att_d, alignment: 'center' },
                      ],
                    ],
                  },
                  // layout: 'lightHorizontalLines'
                },
              ],
            ],
          },
          margin: [0, 0, 0, 15],
        },

        // Part-3 Table
        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: safe(l.sections?.part3),
                  bold: true,
                  fillColor: '#eeeeee',
                  alignment: 'center',
                },
              ],
            ],
          },
        },
        {
          table: {
            headerRows: 1,
            widths: [25, '*', 45, 55, 60, 50, 50, 55],
            body: [
              [
                l.headers?.sl,
                l.headers?.inv_no,
                l.headers?.date,
                l.headers?.supplier,
                l.headers?.bin,
                l.headers?.desc,
                l.headers?.value,
                l.headers?.vat_sd,
              ].map((h) => ({ text: h, bold: true, alignment: 'center', fontSize: 7 })),
              ...items.map((row: any, i: number) => [
                { text: toBnNum(i + 1), alignment: 'center' },
                safe(row.inv_no),
                toBnNum(safe(row.date)),
                safe(row.supplier),
                toBnNum(safe(row.supplier_bin)),
                safe(row.desc),
                { text: toBnNum(safe(row.val)), alignment: 'right' },
                { text: toBnNum(safe(row.vat_sd)), alignment: 'right' },
              ]),
              [
                {
                  text: safe(l.headers?.table_total),
                  colSpan: 7,
                  alignment: 'right',
                  bold: true,
                  fontSize: 7,
                },
                {},
                {},
                {},
                {},
                {},
                {},
                { text: toBnNum(safe(targetData.total_vat)), alignment: 'right', bold: true },
              ],
            ],
          },
          margin: [0, 0, 0, 15],
        },

        // Part-4 Declaration
        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: safe(l.sections?.part4),
                  bold: true,
                  fillColor: 'yellow',
                  alignment: 'center',
                },
              ],
            ],
          },
        },
        // { text: safe(l.declaration?.text), margin: [0, 5, 0, 10] },
        {
          table: {
            widths: [80, '*', 150],
            body: [
              [{ text: l.declaration?.text, colSpan: 3 }, {}, {}],
              [
                { text: l.declaration?.name },
                safe(targetData.formData?.applicant_name),
                { text: '', rowSpan: 4, border: [true, true, true, true] },
              ],
              [{ text: l.declaration?.designation }, safe(targetData.formData?.designation), ''],
              [{ text: l.declaration?.date }, toBnNum(new Date().toLocaleDateString()), ''],
              [{ text: l.declaration?.mobile }, '', ''],
              [
                { text: l.declaration?.email },
                '',
                { text: l.declaration?.signature, alignment: 'center' },
              ],
            ],
          },
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`${l.titles?.form}_${lang}.pdf`);
  }

  exportMushak_18_1(data: any, lang: string) {
    const l = (data.labels?.mushak_18_1 || {}) as any;
    const targetData = data.mushak_18_1_data?.[lang] || {};
    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');
    const toBnNum = (n: any) =>
      lang === 'BN' ? n.toString().replace(/\d/g, (d: any) => '০১২৩৪৫৬৭৮৯'[d]) : n;

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: lang === 'BN' ? 'PlaywriteCU' : 'Nunito', fontSize: 8.5 },
      content: [
        {
          columns: [
            { width: '*', text: '' },
            {
              width: 'auto',
              table: {
                widths: [80],
                body: [[{ text: safe(l.titles?.form), alignment: 'center', bold: true }]],
              },
            },
          ],
        },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        {
          text: safe(l.titles?.name),
          alignment: 'center',
          bold: true,
          fontSize: 10,
          margin: [0, 5, 0, 0],
        },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 7.5, margin: [0, 2, 0, 15] },

        // Part-1: General Information
        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: safe(l.sections?.part1),
                  bold: true,
                  fillColor: '#f4cccc',
                  alignment: 'center',
                },
              ],
            ],
          },
          margin: [0, 0, 0, 0],
        },
        {
          table: {
            widths: [200, 10, '*'],
            body: [
              [{ text: l.info?.bin }, ':', toBnNum(safe(targetData.formData?.bin))],
              [{ text: l.info?.tin }, ':', toBnNum(safe(targetData.formData?.tin))],
              [{ text: l.info?.app_name }, ':', safe(targetData.formData?.app_name)],
              [{ text: l.info?.dob }, ':', toBnNum(safe(targetData.formData?.dob))],
              [{ text: l.info?.nationality }, ':', safe(targetData.formData?.nationality)],
            ],
          },
        },

        // Part-2: Educational Qualification
        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: safe(l.sections?.part2),
                  bold: true,
                  fillColor: '#f4cccc',
                  alignment: 'center',
                },
              ],
            ],
          },
          margin: [0, 0, 0, 0],
        },
        {
          table: {
            widths: [200, 10, '*'],
            body: [
              [{ text: l.info?.last_degree }, ':', safe(targetData.formData?.last_degree)],
              [{ text: l.info?.inst }, ':', safe(targetData.formData?.inst)],
            ],
          },
        },

        // Part-3: Eligibility
        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: safe(l.sections?.part3),
                  bold: true,
                  fillColor: '#f4cccc',
                  alignment: 'center',
                },
              ],
            ],
          },
          margin: [0, 0, 0, 0],
        },
        { text: l.info?.eligible_note, fontSize: 7.5, margin: [0, 2, 0, 5] },
        {
          table: {
            widths: ['*', 25],
            body: [
              [
                l.eligibility?.a,
                {
                  text: targetData.formData?.elig_a ? '✔' : '',
                  alignment: 'center',
                  border: [true, true, true, true],
                },
              ],
              [
                l.eligibility?.b,
                { text: targetData.formData?.elig_b ? '✔' : '', alignment: 'center' },
              ],
              [
                l.eligibility?.c,
                { text: targetData.formData?.elig_c ? '✔' : '', alignment: 'center' },
              ],
              [
                l.eligibility?.d,
                { text: targetData.formData?.elig_d ? '✔' : '', alignment: 'center' },
              ],
            ],
          },
        },

        // Part-4: Necessary Documents
        {
          table: {
            widths: ['*'],
            body: [
              [{ text: l.sections?.part4, bold: true, fillColor: '#f4cccc', alignment: 'center' }],
            ],
          },
          margin: [0, 0, 0, 0],
        },
        {
          table: {
            widths: [200, 10, '*'],
            body: [
              [l.docs?.a, ':', targetData.formData?.doc_a || ''],
              [l.docs?.b, ':', targetData.formData?.doc_b || ''],
              [l.docs?.c, ':', targetData.formData?.doc_c || ''],
              [l.docs?.d, ':', targetData.formData?.doc_d || ''],
              [l.docs?.e, ':', targetData.formData?.doc_e || ''],
              [l.docs?.f, ':', targetData.formData?.doc_f || ''],
            ],
          },
        },

        // Part-5: Declaration
        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: safe(l.sections?.part5),
                  bold: true,
                  fillColor: '#f4cccc',
                  alignment: 'center',
                },
              ],
            ],
          },
          margin: [0, 15, 0, 2],
        },
        { text: l.declaration?.text, margin: [0, 5, 0, 10] },
        {
          columns: [
            {
              width: '*',
              stack: [
                { text: `${l.declaration?.name}: ${safe(targetData.formData?.app_name)}` },
                { text: `${l.declaration?.designation}: ________________` },
              ],
            },
          ],
        },
      ],
    };
    pdfMake.createPdf(docDef).download(`${l.titles?.form}_${lang}.pdf`);
  }
  //Mushak 18.2 here
  exportMushak_18_2(data: any, lang: string) {
    const l = (data.labels?.mushak_18_2 || {}) as any;
    const targetData = data.mushak_18_2_data?.[lang] || {};
    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');
    const toBnNum = (n: any) =>
      lang === 'BN' ? n.toString().replace(/\d/g, (d: any) => '০১২৩৪৫৬৭৮৯'[d]) : n;

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [50, 40, 50, 40],
      defaultStyle: { font: lang === 'BN' ? 'PlaywriteCU' : 'Nunito', fontSize: 9 },
      content: [
        // Top: applicant info center, form number box right
        {
          columns: [
            { width: '*', text: '' },
            {
              width: '*',
              text: safe(l.header?.applicantInfo),
              fontSize: 8,
              alignment: 'center',
            },
            {
              width: '*',
              columns: [
                { width: '*', text: '' },
                {
                  width: 'auto',
                  table: {
                    widths: [70],
                    body: [
                      [
                        {
                          text: safe(l.header?.formNumber),
                          alignment: 'center',
                          bold: true,
                          fontSize: 8,
                        },
                      ],
                    ],
                  },
                  alignment: 'right',
                },
              ],
            },
          ],
          margin: [0, 0, 0, 10],
        },

        // Title
        {
          text: safe(l.titles?.form),
          alignment: 'center',
          bold: true,
          fontSize: 12,
          // decoration: 'underline',
          margin: [0, 0, 0, 8],
        },

        // Rule
        {
          text: safe(l.titles?.rule),
          alignment: 'center',
          fontSize: 8,
          margin: [0, 0, 0, 20],
        },

        // Main Table
        {
          table: {
            widths: [20, '*', 10, '*'],
            body: [
              //Name
              [
                { text: toBnNum(1), alignment: 'center', rowSpan: 3 },
                {
                  stack: [
                    { text: safe(l.table?.rows?.row1_label) },
                    { text: safe(l.table?.rows?.row1a) },
                  ],
                },
                { text: ':', rowSpan: 3 },
                { text: safe(targetData.formData?.applicant_name), rowSpan: 3 },
              ],
              // Row 1b: address
              [{}, { text: safe(l.table?.rows?.row1b) }, {}, {}],
              // Row 1c: BNI
              [{}, { text: safe(l.table?.rows?.row1c) }, {}, {}],

              // Row 2
              [
                { text: toBnNum(2), alignment: 'center' },
                { text: safe(l.table?.rows?.row2) },
                { text: ':' },
                { text: safe(targetData.formData?.document_description) },
              ],

              // Row 3
              [
                { text: toBnNum(3), alignment: 'center' },
                { text: safe(l.table?.rows?.row3) },
                { text: ':' },
                { text: safe(targetData.formData?.document_quantity) },
              ],

              // Row 4
              [
                { text: toBnNum(4), alignment: 'center' },
                { text: safe(l.table?.rows?.row4) },
                { text: ':' },
                { text: safe(targetData.formData?.purpose_of_use) },
              ],

              // Row 5
              [
                { text: toBnNum(5), alignment: 'center' },
                { text: safe(l.table?.rows?.row5) },
                { text: ':' },
                { text: safe(targetData.formData?.fee_challan_info) },
              ],

              // Row 6
              [
                { text: toBnNum(6), alignment: 'center' },
                { text: safe(l.table?.rows?.row6) },
                { text: ':' },
                { text: safe(targetData.formData?.officer_info) },
              ],
            ],
          },
          margin: [0, 0, 0, 20],
        },

        // Declaration title box
        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: safe(l.declaration?.title),
                  bold: true,
                  fillColor: '#d9d9d9',
                  fontSize: 7,
                  margin: [4, 3, 4, 3],
                },
              ],
            ],
          },
          margin: [0, 0, 0, 0],
        },

        // Declaration text
        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: safe(l.declaration?.text),
                  fontSize: 8.5,
                  margin: [4, 5, 4, 10],
                  border: [true, false, true, true],
                },
              ],
            ],
          },
          margin: [0, 0, 0, 20],
        },

        // Declaration fields
        {
          table: {
            widths: [120, 10, '*'],
            body: [
              [
                { text: safe(l.declaration?.signature), border: [false, false, false, false] },
                { text: ':', border: [false, false, false, false] },
                {
                  text: safe(targetData.declaration?.signature),
                  border: [false, false, false, false],
                },
              ],
              [
                { text: safe(l.declaration?.name), border: [false, false, false, false] },
                { text: ':', border: [false, false, false, false] },
                { text: safe(targetData.declaration?.name), border: [false, false, false, false] },
              ],
              [
                { text: safe(l.declaration?.address), border: [false, false, false, false] },
                { text: ':', border: [false, false, false, false] },
                {
                  text: safe(targetData.declaration?.address),
                  border: [false, false, false, false],
                },
              ],
            ],
          },
        },
      ],
    };

    pdfMake.createPdf(docDef).download(`Mushak_18.2_${lang}.pdf`);
  }
  //Mushak 18.3 here
  exportMushak_18_3(data: any, lang: string) {
    const l = (data.labels?.mushak_18_3 || {}) as any;
    const targetData = data.mushak_18_3_data?.[lang] || {};
    const safe = (val: any) => (val !== undefined && val !== null ? val.toString() : '');

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
    };

    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [50, 40, 50, 40],
      defaultStyle: { font: lang === 'BN' ? 'PlaywriteCU' : 'Nunito', fontSize: 9 },
      content: [
        // Top: applicant info left, form number box right
        // Top: applicant info left, form number box right
        {
          columns: [
            { width: '*', text: '' },
            {
              width: '*',
              text: safe(l.header?.applicantInfo),
              fontSize: 8,
              alignment: 'center',
            },
            {
              width: '*',
              columns: [
                { width: '*', text: '' }, // ← pusher
                {
                  width: 'auto',
                  table: {
                    widths: [70],
                    body: [
                      [
                        {
                          text: safe(l.header?.formNumber),
                          alignment: 'center',
                          bold: true,
                          fontSize: 8,
                        },
                      ],
                    ],
                  },
                  alignment: 'right',
                },
              ],
            },
          ],
          margin: [0, 0, 0, 10],
        },

        // Title
        {
          text: safe(l.titles?.form),
          alignment: 'center',
          bold: true,
          fontSize: 12,
          // decoration: 'underline',
          margin: [0, 0, 0, 4],
        },

        // Rule
        {
          text: safe(l.titles?.rule),
          alignment: 'center',
          fontSize: 8,
          margin: [0, 0, 0, 20],
        },

        // Date
        { text: safe(l.body?.date), margin: [0, 0, 0, 15] },

        // To / Bararbar
        { text: safe(l.body?.to), margin: [0, 0, 0, 2] },
        { text: safe(l.body?.designation), margin: [0, 0, 0, 2] },
        { text: safe(l.body?.department) + ' __________', margin: [0, 0, 0, 15] },

        // Subject
        { text: safe(l.body?.subject), bold: true, margin: [0, 0, 0, 10] },

        // Salutation
        { text: safe(l.body?.salutation), margin: [0, 0, 0, 6] },

        // Main paragraph
        {
          text: [
            {
              text: safe(targetData.body?.applicantName),
              // '.............................................'
            },
            { text: ' ' + safe(targetData.body?.paragraph) },
          ],
          margin: [0, 0, 0, 15],
          align: 'justify',
        },

        // Reason intro
        { text: safe(l.body?.reason_intro), margin: [0, 0, 0, 10] },

        // Reasons (a) and (b)
        {
          text:
            safe(targetData.body?.reason_a) ||
            (lang === 'BN' ? '(ক)............................' : '(a)............................'),
          margin: [0, 0, 0, 8],
        },
        {
          text:
            safe(targetData.body?.reason_b) ||
            (lang === 'BN' ? '(খ)............................' : '(b)............................'),
          margin: [0, 0, 0, 20],
        },

        // Closing
        { text: safe(l.body?.closing), margin: [0, 0, 0, 40] },

        // Declaration section box
        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: safe(l.declaration?.title),
                  bold: true,
                  fillColor: '#d9d9d9',
                  fontSize: 7,
                  margin: [4, 3, 4, 3],
                },
              ],
            ],
          },
          margin: [0, 0, 0, 0],
        },
        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  text: safe(l.declaration?.text),
                  fontSize: 8.5,
                  margin: [4, 5, 4, 10],
                  border: [true, false, true, true],
                },
              ],
            ],
          },
          margin: [0, 0, 0, 10],
        },

        // Declaration fields
        {
          table: {
            widths: [120, 10, '*'],
            body: [
              [
                { text: safe(l.declaration?.name), border: [false, false, false, false] },
                { text: ':', border: [false, false, false, false] },
                { text: safe(targetData.declaration?.name), border: [false, false, false, false] },
              ],
              [
                { text: safe(l.declaration?.address), border: [false, false, false, false] },
                { text: ':', border: [false, false, false, false] },
                {
                  text: safe(targetData.declaration?.address),
                  border: [false, false, false, false],
                },
              ],
              [
                { text: safe(l.declaration?.bin), border: [false, false, false, false] },
                { text: ':', border: [false, false, false, false] },
                { text: safe(targetData.declaration?.bin), border: [false, false, false, false] },
              ],
            ],
          },
        },
      ],
    };

    pdfMake.createPdf(docDef).download(`Mushak_18.3_${lang}.pdf`);
  }

  exportMushak_2_1(data: any, lang: string) {
    debugger
    const l = (data.labels?.mushak_2_1 || {}) as any;
    const targetData = data.mushak_2_1_data?.[lang] || {};
    const al = l.address_labels || {};
    const bl = l.branch_labels || {};
    const safe = (val: any) => (val !== undefined && val !== null) ? val.toString() : '';
    const toBnNum = (n: any) => lang === 'BN' ? n.toString().replace(/\d/g, (d: any) => "০১২৩৪৫৬৭৮৯"[d]) : n;

    (pdfMake as any).fonts = {
      Nunito: {
        normal: window.location.origin + '/assets/fonts/Nunito-Regular.ttf',
        bold: window.location.origin + '/assets/fonts/Nunito-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/Nunito-Italic.ttf',
        bolditalics: window.location.origin + '/assets/fonts/Nunito-BoldItalic.ttf',
      },
      PlaywriteCU: {
        normal: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bold: window.location.origin + '/assets/fonts/Kalpurush-Bold.ttf',
        italics: window.location.origin + '/assets/fonts/kalpurush.ttf',
        bolditalics: window.location.origin + '/assets/fonts/kalpurush.ttf',
      },
    };
    const tickImage = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMgAAADICAYAAACtWK6eAAAABmJLR0QA/wD/AP+gvaeTAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAB3RJTUUH5AYWDA8p7zZ7WAAAAB1pVFh0Q29tbWVudAAAAAAAQ3JlYXRlZCB3aXRoIEdJTVBkLm3CYAAAAnZJREFUeNrt17FpAmEUBuDfS8YRLBygkYidpLOInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ2kkYidpJGInaSRiJ+8B9k6YwVwN+f6AAAAAElFTkSuQmCC';
    const docDef: any = {
      pageSize: 'A4',
      pageMargins: [30, 30, 30, 30],
      defaultStyle: { font: lang === 'BN' ? 'PlaywriteCU' : 'Nunito', fontSize: 8 },
      content: [
        // --- PAGE 1: Identity & Address ---
        { columns: [{ width: '*', text: '' }, { width: 'auto', table: { body: [[{ text: safe(l.titles?.form), bold: true }]] } }], margin: [0, 0, 0, 10] },
        { text: safe(l.titles?.gov), alignment: 'center', bold: true },
        { text: safe(l.titles?.nbr), alignment: 'center', bold: true },
        { text: safe(l.titles?.name), alignment: 'center', bold: true, fontSize: 10, margin: [0, 5, 0, 0] },
        { text: safe(l.titles?.rule), alignment: 'center', fontSize: 7, margin: [0, 0, 0, 15] },

        // Sec 1 & 2 [cite: 9-13]
        { text: safe(l.sections?.part1), bold: true, margin: [0, 5, 0, 2] },
        { table: { widths: ['*'], body: [[{ text: toBnNum(safe(targetData.formData?.tin)), margin: [5, 2] }]] } },
        { text: safe(l.sections?.part2), bold: true, margin: [0, 5, 0, 2] },
        { table: { widths: ['*'], body: [[{ text: safe(targetData.formData?.person_name), margin: [5, 2] }]] } },

        // Sec 3: Address Table
        { text: safe(l.sections?.part3), bold: true, margin: [0, 5, 0, 2] },
        {
          table: {
            widths: [87, 100, 100, 100, 100],
            body: [
              [
                { text: al.address, rowSpan: 5, alignment: 'center', margin: [0, 25, 0, 0] },
                { text: al.fill_any_one, colSpan: 4, alignment: 'center', bold: true },
                {}, {}, {}
              ],
              [
                {},
                { text: al.town_address, colSpan: 2, alignment: 'center', fillColor: '#f2f2f2' },
                {},
                { text: al.village_address, colSpan: 2, alignment: 'center', fillColor: '#f2f2f2' }
              ],
              [
                {},
                { text: al.holding, bold: true },
                { text: safe(targetData.formData?.address?.holding) },
                { text: al.mohalla, bold: true },
                { text: safe(targetData.formData?.address?.mohalla) }
              ],
              [
                {},
                { text: al.road, bold: true },
                { text: safe(targetData.formData?.address?.road) },
                { text: al.village, bold: true },
                { text: safe(targetData.formData?.address?.village) }
              ],
              [
                {},
                { text: al.area, bold: true },
                { text: safe(targetData.formData?.address?.area) },
                { text: al.thana, bold: true },
                { text: safe(targetData.formData?.address?.thana) }
              ],
              [
                { text: al.district, bold: true },
                { text: safe(targetData.formData?.address?.district), colSpan: 2 },
                {},
                { text: al.upazila, bold: true },
                { text: safe(targetData.formData?.address?.upazila) }
              ],
              [
                { text: al.post_code, bold: true },
                { text: toBnNum(safe(targetData.formData?.address?.post_code)), colSpan: 2 },
                {},
                { text: al.mouza, bold: true },
                { text: toBnNum(safe(targetData.formData?.address?.mouza)) }
              ],
              [
                { text: al.phone, bold: true },
                { text: safe(targetData.formData?.address?.phone), colSpan: 2 },
                {},
                { text: al.mobile, bold: true },
                { text: safe(targetData.formData?.address?.mobile) }
              ],
              [
                { text: al.email, bold: true },
                { text: safe(targetData.formData?.address?.email), colSpan: 2 },
                {},
                { text: al.fax, bold: true },
                { text: safe(targetData.formData?.address?.fax) }
              ],
              [
                { text: al.web, bold: true },
                { text: safe(targetData.formData?.address?.web), colSpan: 4 }
              ]
            ]
          },
          margin: [0, 0, 0, 10]
        },

        // Section 4: Address of Branch Units
        { text: safe(l.sections?.part4), bold: true, margin: [0, 10, 0, 2] },
        {
          table: {
            widths: [30, 60, '*', 60, '*'],
            body: [
              [
                { text: bl.sl, alignment: 'center', bold: true },
                { text: bl.addr_header, colSpan: 2, alignment: 'center', bold: true },
                {},
                { text: bl.contact_header, colSpan: 2, alignment: 'center', bold: true },
                {}
              ],
              ...(targetData.branchData || []).flatMap((branch: any, index: number) => [
                [
                  { text: toBnNum(index + 1), rowSpan: 3, alignment: 'center', margin: [0, 15, 0, 0] },
                  { text: al.address, bold: true, rowSpan: 2 },
                  { text: safe(branch.address), rowSpan: 2 },
                  { text: al.mobile, bold: true },
                  { text: toBnNum(safe(branch.mobile)) }
                ],
                [
                  {},
                  { text: '', border: [true, false, true, false] },
                  { text: '', border: [true, false, true, false] },
                  { text: al.phone, bold: true },
                  { text: toBnNum(safe(branch.phone)) }
                ],
                [
                  {},
                  { text: al.mouza, bold: true },
                  { text: safe(branch.mouza) },
                  { text: al.email, bold: true },
                  { text: safe(branch.email) }
                ]
              ])
            ]
          }
        },
        { text: bl.note, fontSize: 7, alignment: 'right', margin: [0, 2, 0, 10] },

        { text: safe(l.sections?.part5), bold: true, margin: [0, 10, 0, 2] },
        {
          table: {
            headerRows: 1, widths: [30, '*', '*', '*', '*'],
            body: [
              [l.bank_headers?.sl, l.bank_headers?.acc_name, l.bank_headers?.acc_no, l.bank_headers?.bank_name, l.bank_headers?.branch].map(h => ({ text: h, bold: true, alignment: 'center' })),
              ...(targetData.bankData || []).map((b: any, i: number) => [toBnNum(i + 1), safe(b.acc_name), toBnNum(safe(b.acc_no)), safe(b.bank), safe(b.branch)])
            ]
          },
          pageBreak: 'after' // Correct placement to avoid Error
        },
        { text: bl.note, fontSize: 7, alignment: 'right', margin: [0, 2, 0, 10] },

        // --- Section-6 ---
        { text: safe(l.sections?.part6), bold: true, margin: [0, 5, 0, 2] },
        { table: { widths: [80, '*'], body: [[l.bank_headers?.amount, toBnNum(safe(targetData.formData?.turnover))], [l.bank_headers?.inWord, ' ']] } },

        // --- Section-7 ---
        { text: safe(l.sections?.part7), bold: true, margin: [0, 10, 0, 5] },
        {
          columns: [
            {
              width: 'auto',
              table: {
                widths: [20, 150],
                body: [
                  [
                    { text: '' },
                    { text: safe(l.reg_labels?.vat_reg), margin: [5, 2] }
                  ]
                ]
              },
              margin: [0, 0, 15, 0]
            },
            {
              width: 'auto',
              table: {
                widths: [20, 170],
                body: [
                  [
                    { text: '' },
                    { text: safe(l.reg_labels?.tt_enlist), margin: [5, 2] }
                  ]
                ]
              }
            }
          ]
        },
        // --- Section-8 ---
        {
          columns: [
            { text: safe(l.sections?.part8), bold: true, width: 'auto', margin: [0, 5, 10, 5] },
            {
              width: 'auto',
              table: {
                widths: [20, 60],
                body: [[{ text: '' }, { text: safe(l.withholding?.yes), alignment: 'center' }]]
              }
            },
            {
              width: 'auto',
              table: {
                widths: [20, 60],
                body: [[{ text: '' }, { text: safe(l.withholding?.no), alignment: 'center' }]]
              },
              margin: [10, 0, 0, 0]
            }
          ],
          margin: [0, 10, 0, 5]
        },
        {
          columns: [
            {
              width: 'auto',
              stack: [
                { table: { widths: [15, 100], body: [[{ text: '' }, { text: safe(l.withholding?.govt), margin: [2, 1] }]] }, margin: [0, 0, 0, 3] },
                { table: { widths: [15, 100], body: [[{ text: '' }, { text: safe(l.withholding?.edu), margin: [2, 1] }]] }, margin: [0, 0, 0, 3] }
              ]
            },
            {
              width: 'auto',
              stack: [
                { table: { widths: [15, 120], body: [[{ text: '' }, { text: safe(l.withholding?.ngo), margin: [2, 1] }]] }, margin: [3, 0, 0, 3] },
                { table: { widths: [15, 120], body: [[{ text: '' }, { text: safe(l.withholding?.ltu), margin: [2, 1] }]] }, margin: [3, 0, 0, 3] }
              ]
            },
            {
              width: 'auto',
              stack: [
                { table: { widths: [15, 110], body: [[{ text: '' }, { text: safe(l.withholding?.pub_ltd), margin: [2, 1] }]] }, margin: [3, 0, 0, 3] },
                { table: { widths: [15, 110], body: [[{ text: '' }, { text: safe(l.withholding?.bank), margin: [2, 1] }]] }, margin: [3, 0, 0, 3] },
              ]
            }
          ]
        },

        { text: safe(l.sections?.part9), bold: true, margin: [0, 10, 0, 5] },
        {
          columns: [
            // Column 1
            {
              width: 'auto',
              stack: [
                { table: { widths: [15, 100], body: [[{ text: '' }, { text: safe(l.nature?.natural), margin: [2, 1] }]] }, margin: [0, 0, 0, 3] },
                { table: { widths: [15, 100], body: [[{ text: '' }, { text: safe(l.nature?.pvt_ltd), margin: [2, 1] }]] }, margin: [0, 0, 0, 3] }
              ]
            },
            // Column 2
            {
              width: 'auto',
              stack: [
                { table: { widths: [15, 110], body: [[{ text: '' }, { text: safe(l.nature?.proprietor), margin: [2, 1] }]] }, margin: [0, 0, 0, 3] },
                { table: { widths: [15, 110], body: [[{ text: '' }, { text: safe(l.nature?.pub_ltd), margin: [2, 1] }]] }, margin: [0, 0, 0, 3] }
              ]
            },
            // Column 3
            {
              width: 'auto',
              stack: [
                { table: { widths: [15, 100], body: [[{ text: '' }, { text: safe(l.nature?.partnership), margin: [2, 1] }]] }, margin: [0, 0, 0, 3] },
                { table: { widths: [15, 100], body: [[{ text: '' }, { text: safe(l.nature?.foreign), margin: [2, 1] }]] }, margin: [0, 0, 0, 3] }
              ]
            },
            // Column 4
            {
              width: 'auto',
              stack: [
                { table: { widths: [15, 100], body: [[{ text: '' }, { text: safe(l.nature?.intl_org), margin: [2, 1] }]] }, margin: [0, 0, 0, 3] },
                { table: { widths: [15, 100], body: [[{ text: '' }, { text: safe(l.nature?.diplomat), margin: [2, 1] }]] }, margin: [0, 0, 0, 3] }
              ]
            }
          ],
          columnGap: 5
        },
        {
          margin: [0, 3, 0, 0],
          columns: [
            {
              width: 'auto',
              table: {
                widths: [15, 100],
                body: [[{ text: '' }, { text: safe(l.nature?.others), margin: [2, 1] }]]
              }
            },
            {
              width: '*',
              table: {
                widths: ['*'],
                body: [[{ text: safe(targetData.formData?.nature_others_text), minHeight: 10, margin: [2, 1] }]]
              },
              margin: [5, 0, 0, 0]
            }
          ]
        },

        // --- Section-10 --- 
        { text: safe(l.sections?.part10), bold: true, margin: [0, 10, 0, 5] },
        {
          columns: [
            // Column 1 
            {
              width: 'auto',
              stack: [
                {
                  table: {
                    widths: [15, 110],
                    body: [[
                      { text: '' },
                      { text: safe(l.other_taxes?.sd), margin: [2, 1] }
                    ]]
                  }
                }
              ]
            },
            // Column 2 
            {
              width: 'auto',
              stack: [
                {
                  table: {
                    widths: [15, 110],
                    body: [[
                      { text: '' },
                      { text: safe(l.other_taxes?.excise), margin: [2, 1] }
                    ]]
                  }
                }
              ]
            },
            // Column 3 
            {
              width: 'auto',
              stack: [
                {
                  table: {
                    widths: [15, 110],
                    body: [[
                      { text: '' },
                      { text: safe(l.other_taxes?.surcharge), margin: [2, 1] }
                    ]]
                  }
                }
              ]
            }
          ],
          columnGap: 5
        },
        // --- Section-11 --- 
        { text: safe(l.sections?.part11), bold: true, margin: [0, 10, 0, 2] },
        { table: { widths: [100], body: [[toBnNum(safe(targetData.formData?.effective_date))]] } },

        // --- Section-12 --- 
        { text: safe(l.sections?.part12), bold: true, margin: [0, 10, 0, 5] },
        {
          columns: [
            // Column 1 
            {
              width: 'auto',
              stack: [
                {
                  table: {
                    widths: [15, 110],
                    body: [[
                      { text: '' },
                      { text: safe(l.app_nature?.mandatory), margin: [2, 1] }
                    ]]
                  }
                }
              ]
            },
            // Column 2 
            {
              width: 'auto',
              stack: [
                {
                  table: {
                    widths: [15, 110],
                    body: [[
                      { text: '' },
                      { text: safe(l.app_nature?.voluntary), margin: [2, 1] }
                    ]]
                  }
                }
              ]
            },
            // Column 3 
            {
              width: 'auto',
              stack: [
                {
                  table: {
                    widths: [15, 110],
                    body: [[
                      { text: '' },
                      { text: safe(l.app_nature?.suo_moto), margin: [2, 1] }
                    ]]
                  }
                }
              ]
            }
          ],
          columnGap: 5
        },

        // --- Section-13 --- 
        { text: safe(l.sections?.part13), bold: true, margin: [0, 10, 0, 5] },
        {
          stack: [
            {
              columns: [
                {
                  width: 'auto',
                  table: {
                    widths: [20],
                    body: [[{ text: '', minHeight: 15 }]]
                  }
                },
                { text: safe(l.app_type?.new), margin: [5, 2] }
              ],
              margin: [15, 0, 0, 2]
            },
            {
              columns: [
                {
                  width: 'auto',
                  table: {
                    widths: [20],
                    body: [[{ text: '', minHeight: 15 }]]
                  }
                },
                { text: safe(l.app_type?.re_reg), margin: [5, 2] }
              ],
              margin: [15, 0, 0, 10]
            }
          ]
        },
        { text: safe(l.app_type?.re_reg_bin), fontSize: 8, alignment: 'left', margin: [0, 0, 0, 0] },

        { text: safe(l.app_type?.part13_sub), fontSize: 9, bold: true, margin: [0, 5, 0, 5] },
        {
          columns: [
            {
              width: '49%',
              table: {
                widths: [35, '*'],
                body: [
                  [{ text: safe(l.app_type?.sl), bold: true, alignment: 'center' }, { text: safe(l.app_type?.old_bin), bold: true, alignment: 'center' }],
                  ...Array.from({ length: 10 }, (_, i) => [
                    { text: (i + 1).toLocaleString(lang === 'BN' ? 'bn-BD' : 'en-US') + '.', alignment: 'center' },
                    { text: safe(targetData.formData?.old_bin_list?.[i]), minHeight: 15 }
                  ])
                ]
              }
            },
            {
              width: '49%',
              table: {
                widths: [35, '*'],
                body: [
                  [{ text: safe(l.app_type?.sl), bold: true, alignment: 'center' }, { text: safe(l.app_type?.old_bin), bold: true, alignment: 'center' }],
                  ...Array.from({ length: 10 }, (_, i) => [
                    { text: (i + 11).toLocaleString(lang === 'BN' ? 'bn-BD' : 'en-US') + '.', alignment: 'center' },
                    { text: safe(targetData.formData?.old_bin_list?.[i + 10]), minHeight: 15 }
                  ])
                ]
              },
              margin: [5, 0, 0, 0]
            }
          ]
        },
        { text: safe(l.app_type?.extra_paper), fontSize: 8, alignment: 'right', margin: [0, 3, 0, 0] },

        // --- Section-14 ---
        { text: safe(l.sections?.part14), bold: true, margin: [0, 55, 0, 5] },
        {
          table: {
            headerRows: 1,
            widths: [20, '*', 70, 45, '*'],
            body: [
              [
                { text: safe(l.directors?.sl), bold: true, alignment: 'center' },
                { text: safe(l.directors?.name), bold: true, alignment: 'center' },
                { text: safe(l.directors?.designation), bold: true, alignment: 'center' },
                { text: safe(l.directors?.share), bold: true, alignment: 'center' },
                { text: safe(l.directors?.id_info), bold: true, alignment: 'center' }
              ],
              ...(targetData.directors || Array(6).fill({})).map((d: any, i: number) => [
                { text: (i + 1).toLocaleString(lang === 'BN' ? 'bn-BD' : 'en-US'), alignment: 'center', margin: [0, 15] },
                { text: safe(d.name), margin: [2, 15] },
                { text: safe(d.designation), margin: [2, 15] },
                { text: (d.share ? d.share.toLocaleString(lang === 'BN' ? 'bn-BD' : 'en-US') : ''), alignment: 'center', margin: [0, 15] },
                {
                  table: {
                    widths: [75, '*'],
                    body: [
                      [{ text: safe(l.directors?.id_type), margin: [2, 1] }, { text: safe(l.directors?.nid_tin), margin: [2, 1] }],
                      [{ text: safe(l.directors?.nid), margin: [2, 1] }, { text: safe(d.nid), margin: [2, 1] }],
                      [{ text: safe(l.directors?.passport), margin: [2, 1] }, { text: safe(d.passport), margin: [2, 1] }],
                      [{ text: safe(l.directors?.issue_country), margin: [2, 1] }, { text: safe(d.issue_country), margin: [2, 1] }]
                    ]
                  },
                  layout: {
                    hLineWidth: (i: number, node: any): number => (i === 0 || i === node.table.body.length) ? 0 : 0.5,
                    vLineWidth: (i: number, node: any): number => (i === 1) ? 0.5 : 0,
                    hLineColor: (): string => '#000000',
                    vLineColor: (): string => '#000000'
                  }
                }
              ])
            ]
          }
        },
        { text: safe(l.directors?.extra_paper), fontSize: 8, alignment: 'right', margin: [0, 3, 0, 0] },

        // --- Section-15 ---
        { text: safe(l.sections?.part15), bold: true, margin: [0, 10, 0, 5] },
        {
          columns: [
            // Column 1
            {
              width: 'auto',
              stack: [
                { table: { widths: [15, 110], body: [[{ text: '' }, { text: safe(l.biz_nature?.importer), margin: [2, 1] }]] }, margin: [0, 0, 0, 2] },
                { table: { widths: [15, 110], body: [[{ text: '' }, { text: safe(l.biz_nature?.supplier_mfg), margin: [2, 1] }]] }, margin: [0, 0, 0, 2] },
                { table: { widths: [15, 110], body: [[{ text: '' }, { text: safe(l.biz_nature?.agri_fish), margin: [2, 1] }]] }, margin: [0, 0, 0, 2] },
                { table: { widths: [15, 110], body: [[{ text: '' }, { text: safe(l.biz_nature?.others), margin: [2, 1] }]] } }
              ]
            },
            // Column 2
            {
              width: 'auto',
              stack: [
                { table: { widths: [15, 115], body: [[{ text: '' }, { text: safe(l.biz_nature?.service), margin: [2, 1] }]] }, margin: [0, 0, 0, 2] },
                { table: { widths: [15, 115], body: [[{ text: '' }, { text: safe(l.biz_nature?.mineral), margin: [2, 1] }]] }, margin: [0, 0, 0, 2] },
                {
                  table: {
                    widths: [140],
                    body: [[{ text: safe(targetData.formData?.biz_nature_others_text), minHeight: 15 }]]
                  },
                }
              ],
            },
            // Column 3
            {
              width: 'auto',
              stack: [
                { table: { widths: [15, 110], body: [[{ text: '' }, { text: safe(l.biz_nature?.exporter), margin: [2, 1] }]] }, margin: [0, 0, 0, 2] },
                { table: { widths: [15, 110], body: [[{ text: '' }, { text: safe(l.biz_nature?.supplier_comm), margin: [2, 1] }]] }, margin: [0, 0, 0, 2] }
              ],
              margin: [1, 0, 0, 0]
            }
          ],
          columnGap: 5
        },

        // --- Section-16 ---
        { text: safe(l.sections?.part16), bold: true, margin: [0, 10, 0, 5] },

        { text: safe(l.eco_nature?.part16_ka), margin: [0, 0, 0, 5] },
        {
          columns: [
            // Column 1
            {
              width: 'auto',
              stack: [
                { table: { widths: [15, 110], body: [[{ text: targetData.formData?.is_retail ? '√' : '' }, { text: safe(l.eco_nature?.retail), margin: [2, 1] }]] }, margin: [0, 0, 0, 2] },
                { table: { widths: [15, 110], body: [[{ text: targetData.formData?.is_construction ? '√' : '' }, { text: safe(l.eco_nature?.construction), margin: [2, 1] }]] }, margin: [0, 0, 0, 2] },
                { table: { widths: [15, 110], body: [[{ text: targetData.formData?.is_mineral ? '√' : '' }, { text: safe(l.eco_nature?.mineral), margin: [2, 1] }]] } }
              ],
              margin: [0, 0, 0, 0]
            },
            // Column 2
            {
              width: 'auto',
              stack: [
                { table: { widths: [15, 110], body: [[{ text: targetData.formData?.is_wholesale ? '√' : '' }, { text: safe(l.eco_nature?.wholesale), margin: [2, 1] }]] }, margin: [0, 0, 0, 2] },
                { table: { widths: [15, 110], body: [[{ text: targetData.formData?.is_seasonal ? '√' : '' }, { text: safe(l.eco_nature?.seasonal), margin: [2, 1] }]] }, margin: [0, 0, 0, 2] },
                { table: { widths: [15, 110], body: [[{ text: targetData.formData?.is_agri_fish ? '√' : '' }, { text: safe(l.eco_nature?.agri_fish), margin: [2, 1] }]] } }
              ],
              margin: [0, 0, 0, 0]
            },
            // Column 3
            {
              width: 'auto',
              stack: [
                { table: { widths: [15, 110], body: [[{ text: targetData.formData?.is_mfg ? '√' : '' }, { text: safe(l.eco_nature?.mfg), margin: [2, 1] }]] }, margin: [0, 0, 0, 2] },
                { table: { widths: [15, 110], body: [[{ text: targetData.formData?.is_service ? '√' : '' }, { text: safe(l.eco_nature?.service), margin: [2, 1] }]] }, margin: [0, 0, 0, 2] },
                { table: { widths: [15, 110], body: [[{ text: targetData.formData?.is_eco_others ? '√' : '' }, { text: safe(l.eco_nature?.others), margin: [2, 1] }]] } }
              ],
              margin: [0, 0, 0, 0]
            }
          ]
        },
        { table: { widths: [422], body: [[{ text: safe(targetData.formData?.eco_others_text), minHeight: 20 }]] }, margin: [0, 5, 20, 10] },

        {
          table: {
            widths: ['*'],
            body: [
              [
                {
                  stack: [
                    { text: safe(l.eco_nature?.part16_kha), bold: true, fontSize: 10 },
                    { text: '', fontSize: 9, margin: [0, 2, 0, 0] }
                  ],
                  fillColor: '#f9f9f9'
                }
              ]
            ]
          },
          margin: [0, 5, 20, 10]
        },
        { table: { widths: ['*'], body: [[{ text: safe(targetData.formData?.eco_others_text), minHeight: 20 }]] }, margin: [0, 5, 20, 10] },

         // --- Section-17 ---
        { text: safe(l.sections?.part17), bold: true, margin: [0, 10, 0, 5] },
        {
          table: {
            widths: ['*'],
            body: [[
              {
                stack: [
                  { text: safe(l.signatory_type?.signatory_header), margin: [0, 5, 0, 10] },

                  {
                    columns: [
                      {
                        width: '50%',
                        stack: [
                          { columns: [{ table: { widths: [15], body: [[{ text: targetData.formData?.signatory_type === 'owner' ? '√' : '', minHeight: 12 }]] }, width: 'auto' }, { text: safe(l.signatory_type?.owner), margin: [5, 2] }], margin: [0, 2] },
                          { columns: [{ table: { widths: [15], body: [[{ text: targetData.formData?.signatory_type === 'partner' ? '√' : '', minHeight: 12 }]] }, width: 'auto' }, { text: safe(l.signatory_type?.partner), margin: [5, 2] }], margin: [0, 2] },
                          { columns: [{ table: { widths: [15], body: [[{ text: targetData.formData?.signatory_type === 'others' ? '√' : '', minHeight: 12 }]] }, width: 'auto' }, { text: safe(l.signatory_type?.others), margin: [5, 2] }], margin: [0, 2] }
                        ]
                      },
                      {
                        width: '50%',
                        stack: [
                          { columns: [{ table: { widths: [15], body: [[{ text: targetData.formData?.signatory_type === 'director' ? '√' : '', minHeight: 12 }]] }, width: 'auto' }, { text: safe(l.signatory_type?.director), margin: [5, 2] }], margin: [0, 2] },
                          { columns: [{ table: { widths: [15], body: [[{ text: targetData.formData?.signatory_type === 'officer' ? '√' : '', minHeight: 12 }]] }, width: 'auto' }, { text: safe(l.signatory_type?.officer), margin: [5, 2] }], margin: [0, 2] },
                          { table: { widths: ['*'], body: [[{ text: safe(targetData.formData?.signatory_others_text), minHeight: 15 }]] }, margin: [0, 2] }
                        ]
                      }
                    ]
                  },

                  {
                    columns: [
                      { text: safe(l.signatory_type?.first_name), width: 'auto', margin: [0, 10, 5, 0] },
                      { table: { widths: ['*'], body: [[{ text: safe(targetData.formData?.sig_first_name), minHeight: 15 }]] }, margin: [0, 8, 10, 0] },
                      { text: safe(l.signatory_type?.last_name), width: 'auto', margin: [0, 10, 5, 0] },
                      { table: { widths: ['*'], body: [[{ text: safe(targetData.formData?.sig_last_name), minHeight: 15 }]] }, margin: [0, 8, 0, 0] }
                    ]
                  },

                  { text: safe(l.signatory_type?.id_info), margin: [0, 10, 0, 5], bold: true },
                  {
                    table: {
                      widths: ['*', 50, '*'],
                      body: [
                        [{ text: safe(l.signatory_type?.passport), bold: true, alignment: 'center' }, 
                          { text: safe(l.signatory_type?.or), alignment: 'center', border: [false, true, false, true] }, 
                          { text: safe(l.signatory_type?.nid), bold: true, alignment: 'center' }],
                        [
                          {
                            table: {
                              widths: [80, '*'],
                              body: [
                                [{ text: safe(l.signatory_type?.number) }, { text: safe(targetData.formData?.sig_passport_no) }],
                                [{ text: safe(l.signatory_type?.issue_country) }, { text: safe(targetData.formData?.sig_passport_country) }],
                                [{ text: safe(l.signatory_type?.issue_date) }, { text: safe(targetData.formData?.sig_passport_issue) }],
                                [{ text: safe(l.signatory_type?.expiry_date) }, { text: safe(targetData.formData?.sig_passport_expiry) }]
                              ]
                            },
                            layout: 'noBorders', rowSpan: 4
                          },
                          {},
                          {
                            table: {
                              widths: [60, '*'],
                              body: [[{ text: safe(l.signatory_type?.number) }, { text: safe(targetData.formData?.sig_nid_no) }]]
                            },
                            layout: 'noBorders'
                          }
                        ],
                        ['', '', ''], ['', '', ''], ['', '', '']  
                      ]
                    }
                  },

                  { text: safe(l.signatory_type?.declaration_text), margin: [0, 20, 0, 30] },

                  {
                    columns: [
                      { text: safe(l.signatory_type?.date) + ' ' + (targetData.formData?.application_date || ''), width: 'auto' },
                      { text: safe(l.signatory_type?.signature), alignment: 'right' }
                    ]
                  }
                ],
                margin: [10, 10, 10, 10]
              }
            ]]
          }
        }
      ]
    };
    pdfMake.createPdf(docDef).download(`${l.titles?.form}_${lang}.pdf`);
  }
}
