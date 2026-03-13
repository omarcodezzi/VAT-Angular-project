import { ComponentFixture, TestBed } from '@angular/core/testing';

import { MushakListComponent } from './mushak-list.component';

describe('MushakListComponent', () => {
  let component: MushakListComponent;
  let fixture: ComponentFixture<MushakListComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [MushakListComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(MushakListComponent);
    component = fixture.componentInstance;                  
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
