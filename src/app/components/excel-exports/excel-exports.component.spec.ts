import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExcelExportsComponent } from './excel-exports.component';

describe('ExcelExportsComponent', () => {
  let component: ExcelExportsComponent;
  let fixture: ComponentFixture<ExcelExportsComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ ExcelExportsComponent ]
    })
    .compileComponents();

    fixture = TestBed.createComponent(ExcelExportsComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
