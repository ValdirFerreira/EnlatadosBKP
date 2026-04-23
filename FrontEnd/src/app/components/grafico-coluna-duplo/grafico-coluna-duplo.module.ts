import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { SelectImageModule } from '../select-image/select-image.module';
import { TranslateModule } from '@ngx-translate/core';
import { GraficoColunaDuploComponent } from './grafico-coluna-duplo.component';


@NgModule({
  declarations: [
    GraficoColunaDuploComponent
  ],
  imports: [
    CommonModule,
    SelectImageModule,
    TranslateModule,
  ],
  exports: [GraficoColunaDuploComponent]
})

export class GraficoColunaDuploModule { }
