import { Routes } from '@angular/router';
import { MushakListComponent } from './mushak-list/mushak-list.component';

export const routes: Routes = [
    { path: '', redirectTo: 'mushak-list', pathMatch: 'full' },
    { path: 'mushak-list', component: MushakListComponent },
];
