import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FiltroGlobalModule } from '../filtroGlobal/filtro-global.module';
import { TranslateModule } from '@ngx-translate/core';
import { SelectImage2Component } from './select-image-2.component';
import { NgSelectModule } from '@ng-select/ng-select';
import { FormsModule } from '@angular/forms';


@NgModule({
  declarations: [
    SelectImage2Component,
    
  ],
  imports: [
    CommonModule,
    FiltroGlobalModule,
    TranslateModule,
    NgSelectModule,
    FormsModule ,
  ],
  exports: [
    SelectImage2Component
  ]
})
export class SelectImage2Module { }
