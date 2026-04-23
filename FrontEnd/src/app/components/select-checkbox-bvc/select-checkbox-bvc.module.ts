import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FiltroGlobalModule } from '../filtroGlobal/filtro-global.module';
import { TranslateModule } from '@ngx-translate/core';
import { NgSelectModule } from '@ng-select/ng-select';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { MatCheckboxModule } from '@angular/material/checkbox';
import { SelectCheckboxBVCComponent } from './select-checkbox-bvc.component';

import { MatListModule } from '@angular/material/list';
import { MatButtonModule } from '@angular/material/button';

@NgModule({
  declarations: [
    SelectCheckboxBVCComponent,
    
  ],
  imports: [
    CommonModule,
    FiltroGlobalModule,
    TranslateModule,
    NgSelectModule,
    MatCheckboxModule,
    FormsModule,
    ReactiveFormsModule,
    MatListModule,
    MatButtonModule
  ],
  exports: [
    SelectCheckboxBVCComponent
  ]
})
export class SelectCheckboxBVCModule { }
