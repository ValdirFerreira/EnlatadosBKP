import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { NgSelectModule } from '@ng-select/ng-select';
import { GraficoColunasEvolutivoMarcasDenominatorsComponent } from './grafico-colunas-evolutivo-marcas-denominators.component';
import { FormsModule } from '@angular/forms';


@NgModule({
  declarations: [
    GraficoColunasEvolutivoMarcasDenominatorsComponent
  ],
  imports: [
    CommonModule,
    NgSelectModule,
    FormsModule ,
  ],
  exports: [GraficoColunasEvolutivoMarcasDenominatorsComponent]
})

export class GraficoColunasEvolutivoMarcasDenominatorsModule { }
