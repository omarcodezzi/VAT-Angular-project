import { CommonModule } from '@angular/common';
import { Component, signal } from '@angular/core'; // Ensure signal is imported
import { RouterLink } from '@angular/router';

@Component({
  selector: 'app-mushak-list',
  standalone: true,
  imports: [CommonModule, RouterLink],
  templateUrl: './mushak-list.component.html',
  styleUrl: './mushak-list.component.css',
})
export class MushakListComponent {
  // Wrap your data in a signal()
  mushakForms = signal([
    { id: 1, title: 'Mushak 6.1', description: 'Purchase book', type: 'INPUT', typeClass: 'bg-warning-light text-warning', selectedLang: 'EN' },
    { id: 2, title: 'Mushak 6.2', description: 'Sales book', type: 'SALES', typeClass: 'bg-primary-light text-primary', selectedLang: 'EN' },
    { id: 3, title: 'Mushak 6.3', description: 'Tax invoice', type: 'INVOICE', typeClass: 'bg-info-light text-info', selectedLang: 'EN' }
  ]);

  download(form: any) {
    console.log('Downloading:', form.title);
  }
}