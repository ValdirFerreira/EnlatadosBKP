import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FiltroGlobalModule } from '../filtroGlobal/filtro-global.module';
import { TranslateModule } from '@ngx-translate/core';
import { NgSelectModule } from '@ng-select/ng-select';
import { FormsModule } from '@angular/forms';
import { SelectCheckboxComponent } from './select-checkbox.component';
import { MatCheckboxModule } from '@angular/material/checkbox';


@NgModule({
  declarations: [
    SelectCheckboxComponent,
    
  ],
  imports: [
    CommonModule,
    FiltroGlobalModule,
    TranslateModule,
    NgSelectModule,
    MatCheckboxModule,
    FormsModule,
  ],
  exports: [
    SelectCheckboxComponent
  ]
})
export class SelectCheckboxModule { }
