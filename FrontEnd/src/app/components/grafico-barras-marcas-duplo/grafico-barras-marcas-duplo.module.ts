import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';

import { SelectImageModule } from '../select-image/select-image.module';
import { GraficoBarrasMarcasDuploComponent } from './grafico-barras-marcas-duplo.component';
import { NgSelectModule } from '@ng-select/ng-select';
import { FormsModule } from '@angular/forms';


@NgModule({
  declarations: [
    GraficoBarrasMarcasDuploComponent
  ],
  imports: [
    CommonModule,
    SelectImageModule,
    NgSelectModule,
    FormsModule ,
  ],
  exports: [GraficoBarrasMarcasDuploComponent]
})

export class GraficoBarrasMarcasDuploModule { }
