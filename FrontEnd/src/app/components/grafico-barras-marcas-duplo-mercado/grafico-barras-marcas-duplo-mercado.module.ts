import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';

import { SelectImageModule } from '../select-image/select-image.module';
import { GraficoBarrasMarcasDuploMercadoComponent } from './grafico-barras-marcas-duplo-mercado.component';
import { NgSelectModule } from '@ng-select/ng-select';
import { FormsModule } from '@angular/forms';


@NgModule({
  declarations: [
    GraficoBarrasMarcasDuploMercadoComponent
  ],
  imports: [
    CommonModule,
    SelectImageModule,
    NgSelectModule,
    FormsModule ,
  ],
  exports: [GraficoBarrasMarcasDuploMercadoComponent]
})

export class GraficoBarrasMarcasDuploMercadoModule { }
