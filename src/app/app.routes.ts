import { Routes } from '@angular/router';
import { MushakListComponent } from './mushak-list/mushak-list.component';
import { ExcelImportComponent } from './excel-import/excel-import.component';

export const routes: Routes = [
  { path: 'dashboard', component: ExcelImportComponent },
  { path: 'mushak-list', component: MushakListComponent },
  { path: '', redirectTo: 'dashboard', pathMatch: 'full' },
  { path: '**', redirectTo: 'dashboard' }
];
