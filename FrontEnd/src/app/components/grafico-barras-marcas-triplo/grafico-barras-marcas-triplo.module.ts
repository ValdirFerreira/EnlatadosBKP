import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';

import { SelectImageModule } from '../select-image/select-image.module';
import { GraficoBarrasMarcasTriploComponent } from './grafico-barras-marcas-triplo.component';
import { NgSelectModule } from '@ng-select/ng-select';
import { FormsModule } from '@angular/forms';


@NgModule({
  declarations: [
    GraficoBarrasMarcasTriploComponent
  ],
  imports: [
    CommonModule,
    SelectImageModule,
    NgSelectModule,
    FormsModule ,
  ],
  exports: [GraficoBarrasMarcasTriploComponent]
})

export class GraficoBarrasMarcasTriploModule { }
