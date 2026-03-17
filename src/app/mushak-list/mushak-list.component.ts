import { CommonModule } from '@angular/common';
import { Component, signal } from '@angular/core'; // Ensure signal is imported
import { RouterLink } from '@angular/router';
import { ExportService } from '../Services/export.service';

@Component({
  selector: 'app-mushak-list',
  standalone: true,
  imports: [CommonModule, RouterLink],
  templateUrl: './mushak-list.component.html',
  styleUrl: './mushak-list.component.css',
})
export class MushakListComponent {
  constructor(
    private exportService: ExportService,
  ) { }
  // Wrap your data in a signal()
  mushakForms = signal([
    { id: 1, title: 'Mushak-2.1', description: 'VAT/Turnover Tax Registration Form', typeClass: 'bg-warning-light text-warning', selectedLang: 'EN' },
    { id: 2, title: 'Mushak-2.3', description: 'Value Added Tax (VAT) Registration Certificate', typeClass: 'bg-warning-light text-warning', selectedLang: 'EN' },
    { id: 3, title: 'Mushak-4.3', description: 'Input-Output Coefficient Declaration', typeClass: 'bg-warning-light text-warning', selectedLang: 'EN' },
    { id: 4, title: 'Mushak-6.1', description: 'Purchase Account book', typeClass: 'bg-warning-light text-warning', selectedLang: 'EN' },
    { id: 5, title: 'Mushak-6.2', description: 'Sales Account book', typeClass: 'bg-warning-light text-warning', selectedLang: 'EN' },
    { id: 6, title: 'Mushak-6.2.1', description: 'Purchase-Sales Account Book', typeClass: 'bg-warning-light text-warning', selectedLang: 'EN' },
    { id: 7, title: 'Mushak-6.3', description: 'Tax Invoice', typeClass: 'bg-primary-light text-primary', selectedLang: 'EN' },
    { id: 8, title: 'Mushak-6.4', description: 'Chalanpatra of Contract-based Production',typeClass: 'bg-info-light text-info', selectedLang: 'EN' },
    { id: 9, title: 'Mushak-6.5', description: 'Goods Transfer Chalan of Centrally Registered Institution',typeClass: 'bg-info-light text-info', selectedLang: 'EN' },
    { id: 10, title: 'Mushak-6.6', description: 'Certificate of Tax Deduction at Source',typeClass: 'bg-info-light text-info', selectedLang: 'EN' },
    { id: 11, title: 'Mushak-6.7', description: 'Credit Note',typeClass: 'bg-info-light text-info', selectedLang: 'EN' },
    { id: 12, title: 'Mushak-6.8', description: 'Debit Note',typeClass: 'bg-info-light text-info', selectedLang: 'EN' },
    { id: 13, title: 'Mushak-6.9', description: 'Turnover Tax Invoice',typeClass: 'bg-info-light text-info', selectedLang: 'EN' },
    { id: 14, title: 'Mushak-6.10', description: 'Information of Purchase-Sales invoices valued above BDT 2 (Two) Lac',typeClass: 'bg-info-light text-info', selectedLang: 'EN' },
    { id: 15, title: 'Mushak-9.1', description: 'Value Added Tax Return Form',typeClass: 'bg-info-light text-info', selectedLang: 'EN' },
    { id: 16, title: 'Mushak-10.1', description: 'Tax Refund Application for Diplomatic or International Organizations',typeClass: 'bg-info-light text-info', selectedLang: 'EN' },
    { id: 17, title: 'Mushak-18.1', description: 'Application for VAT Consultant License',typeClass: 'bg-info-light text-info', selectedLang: 'EN' },
    { id: 18, title: 'Mushak-18.2', description: 'Application for Obtaining Certified Copy of Documents',typeClass: 'bg-info-light text-info', selectedLang: 'EN' },
    { id: 19, title: 'Mushak-18.3', description: 'Application for Obtaining VAT Certificate',typeClass: 'bg-info-light text-info', selectedLang: 'EN' }
  ]);

  selectLanguage(id: number, lang: string) {
    this.mushakForms.update(forms =>
      forms.map(f => f.id === id ? { ...f, selectedLang: lang } : f)
    );
  }

  download(form: any) {
    const lang = form.selectedLang === 'EN' ? 'EN' : 'BN';
    const mushak = form.title;
    const apiEndpoint = 'http://localhost:3000/mushak_values';

    this.exportService.getMergedMushakData(apiEndpoint, lang).subscribe({
      next: (data) => {
        if (lang === 'EN') {
          if (mushak === 'Mushak-2.1') {
            this.exportService.exportMushak_2_1(data, lang);
          }
          else if (mushak === 'Mushak-2.3') {
            this.exportService.exportMushak_2_3(data, lang);
          }
          else if (mushak === 'Mushak-4.3') {
            this.exportService.exportInputOutputCoefficientEnglish(data, lang);
          }
          else if (mushak === 'Mushak-6.1') {
            this.exportService.exportmushak_6_1_English(data, lang);
          }
          else if (mushak === 'Mushak-6.2') {
            this.exportService.exportMushak_6_2_English(data, lang);
          }
          else if (mushak === 'Mushak-6.2.1') {
            this.exportService.exportMushak_6_2_1_English(data, lang);
          }
          else if (mushak === 'Mushak-6.3') {
            this.exportService.exportMushak_6_3_English(data, lang);
          }
          else if (mushak === 'Mushak-6.4') {
            this.exportService.exportMushak_6_4_English(data, lang);
          }
          else if (mushak === 'Mushak-6.5') {
            this.exportService.exportMushak_6_5_English(data, lang);
          }
          else if (mushak === 'Mushak-6.6') {
            this.exportService.exportMushak_6_6_English(data, lang);
          }
          else if (mushak === 'Mushak-6.7') {
            this.exportService.exportMushak_6_7_English(data, lang);
          }
          else if (mushak === 'Mushak-6.8') {
            this.exportService.exportMushak_6_8_English(data, lang);
          }
          else if (mushak === 'Mushak-6.9') {
            this.exportService.exportMushak_6_9_English(data, lang);
          }
          else if (mushak === 'Mushak-6.10') {
            this.exportService.exportMushak_6_10_English(data, lang);
          }
          else if (mushak === 'Mushak-9.1') {
            this.exportService.exportFullMushakPdf(data, lang);
          }
          else if (mushak === 'Mushak-10.1') {
            this.exportService.exportMushak_10_1(data, lang);
          }
          else if (mushak === 'Mushak-18.1') {
            this.exportService.exportMushak_18_1(data, lang);
          }
          else if (mushak === 'Mushak-18.2') {
            this.exportService.exportMushak_18_2(data, lang);
          }
          else {
            this.exportService.exportMushak_18_3(data, lang);
          }
        }
        else {
          if (mushak === 'Mushak-2.1') {
            this.exportService.exportMushak_2_1(data, lang);
          }
          else if (mushak === 'Mushak-2.3') {
            this.exportService.exportMushak_2_3(data, lang);
          }
          else if (mushak === 'Mushak-4.3') {
            this.exportService.exportInputOutputCoefficientBangla(data, lang);
          }
          else if (mushak === 'Mushak-6.1') {
            this.exportService.exportmushak_6_1_Bangla(data, lang);
          }
          else if (mushak === 'Mushak-6.2') {
            this.exportService.exportMushak_6_2_Bangla(data, lang);
          }
          else if (mushak === 'Mushak-6.2.1') {
            this.exportService.exportMushak_6_2_1_Bangla(data, lang);
          }
          else if (mushak === 'Mushak-6.3') {
            this.exportService.exportMushak_6_3_Bangla(data, lang);
          }
          else if (mushak === 'Mushak-6.4') {
            this.exportService.exportMushak_6_4_Bangla(data, lang);
          }
          else if (mushak === 'Mushak-6.5') {
            this.exportService.exportMushak_6_5_Bangla(data, lang);
          }
          else if (mushak === 'Mushak-6.6') {
            this.exportService.exportMushak_6_6_Bangla(data, lang);
          }
          else if (mushak === 'Mushak-6.7') {
            this.exportService.exportMushak_6_7_Bangla(data, lang);
          }
          else if (mushak === 'Mushak-6.8') {
            this.exportService.exportMushak_6_8_Bangla(data, lang);
          }
          else if (mushak === 'Mushak-6.9') {
            this.exportService.exportMushak_6_9_Bangla(data, lang);
          }
          else if (mushak === 'Mushak-6.10') {
            this.exportService.exportMushak_6_10_Bangla(data, lang);
          }
          else if (mushak === 'Mushak-9.1') {
            this.exportService.exportFullMushakPdfBangla(data, lang);
          }
          else if (mushak === 'Mushak-10.1') {
            this.exportService.exportMushak_10_1(data, lang);
          }
          else if (mushak === 'Mushak-18.1') {
            this.exportService.exportMushak_18_1(data, lang);
          }
          else if (mushak === 'Mushak-18.2') {
            this.exportService.exportMushak_18_2(data, lang);
          }
          else{
            this.exportService.exportMushak_18_3(data, lang);
          }
        }
      },
      error: (err) => console.error('API Connection Failed!', err),
    });
  }
}