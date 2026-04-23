import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FiltroGlobalModule } from '../filtroGlobal/filtro-global.module';
import { TranslateModule } from '@ngx-translate/core';
import { SelectImageComponent } from './select-image.component';
import { NgSelectModule } from '@ng-select/ng-select';
import { FormsModule } from '@angular/forms';


@NgModule({
  declarations: [
    SelectImageComponent,
    
  ],
  imports: [
    CommonModule,
    FiltroGlobalModule,
    TranslateModule,
    NgSelectModule,
    FormsModule ,
  ],
  exports: [
    SelectImageComponent
  ]
})
export class SelectImageModule { }
