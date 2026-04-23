import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { NgSelectModule } from '@ng-select/ng-select';
import { GraficoColunasEvolutivoMarcas4RowComponent } from './grafico-colunas-evolutivo-marcas-4-row.component';
import { FormsModule } from '@angular/forms';


@NgModule({
  declarations: [
    GraficoColunasEvolutivoMarcas4RowComponent
  ],
  imports: [
    CommonModule,
    NgSelectModule,
    FormsModule ,
  ],
  exports: [GraficoColunasEvolutivoMarcas4RowComponent]
})

export class GraficoColunasEvolutivoMarcas4RowModule { }
