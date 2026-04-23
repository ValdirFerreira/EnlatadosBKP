import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { NgSelectModule } from '@ng-select/ng-select';
import { GraficoColunasEvolutivoMarcasComponent } from './grafico-colunas-evolutivo-marcas.component';
import { FormsModule } from '@angular/forms';


@NgModule({
  declarations: [
    GraficoColunasEvolutivoMarcasComponent
  ],
  imports: [
    CommonModule,
    NgSelectModule,
    FormsModule ,
  ],
  exports: [GraficoColunasEvolutivoMarcasComponent]
})

export class GraficoColunasEvolutivoMarcasModule { }
