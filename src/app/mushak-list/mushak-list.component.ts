import { Component, computed, ElementRef, signal, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';
import { bestSuggestion, normalizeHeader } from '../Services/string-utils';
import { ExcelService } from '../Services/excel.service';
import { FinalRow, HeaderCheckRow, REQUIRED_HEADERS, TaxHeader } from '../Services/types';
import * as XLSX from 'xlsx';
import { HttpClient } from '@angular/common/http';
import { ExportService } from '../Services/export.service';
import { MushakService } from '../Services/mushak.service';

@Component({
  selector: 'app-excel-import',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './mushak-list.component.html',
  styleUrl: './mushak-list.component.css',
})
export class MushakListComponent {
  mushakForms = [
    {
      id: 1, title: 'Mushak 6.1', description: 'Purchase book for registered person',
      type: 'INPUT', typeClass: 'bg-warning-light text-warning', selectedLang: 'EN'
    },
    {
      id: 2, title: 'Mushak 6.2', description: 'Sales book for registered person',
      type: 'SALES', typeClass: 'bg-primary-light text-primary', selectedLang: 'EN'
    },
    {
      id: 3, title: 'Mushak 6.3', description: 'Tax invoice for supply of goods and services',
      type: 'INVOICE', typeClass: 'bg-info-light text-info', selectedLang: 'EN'
    }
  ];

  constructor(
    private excelService: ExcelService,
    private exportService: ExportService,
    private mushakService: MushakService
  ) { }

  download(form: any) {
    console.log('Downloading:', form.title, 'in', form.selectedLang);
  }
}
