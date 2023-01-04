import { NgModule, NO_ERRORS_SCHEMA } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { NgbModule } from '@ng-bootstrap/ng-bootstrap';
import { ExcelExportsComponent } from './components/excel-exports/excel-exports.component';
import { FormsModule } from '@angular/forms';

@NgModule({
  declarations: [
    AppComponent,
    ExcelExportsComponent
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    NgbModule,FormsModule
  ],
  providers: [],
  bootstrap: [AppComponent],
  schemas: [
    // CUSTOM_ELEMENTS_SCHEMA,
    NO_ERRORS_SCHEMA
  ]
})
export class AppModule { }
