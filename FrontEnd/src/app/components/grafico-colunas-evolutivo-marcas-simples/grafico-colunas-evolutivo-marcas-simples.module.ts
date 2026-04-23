import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { NgSelectModule } from '@ng-select/ng-select';
import { GraficoColunasEvolutivoMarcasSimplesComponent } from './grafico-colunas-evolutivo-marcas-simples.component';
import { FormsModule } from '@angular/forms';


@NgModule({
  declarations: [
    GraficoColunasEvolutivoMarcasSimplesComponent
  ],
  imports: [
    CommonModule,
    NgSelectModule,
    FormsModule ,
  ],
  exports: [GraficoColunasEvolutivoMarcasSimplesComponent]
})

export class GraficoColunasEvolutivoMarcasSimplesModule { }
