import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { SelectImageModule } from '../select-image/select-image.module';
import { GraficoBarrasEvolutivoMarcasComponent } from './grafico-barras-evolutivo-marcas.component';
import { NgSelectModule } from '@ng-select/ng-select';
import { FormsModule } from '@angular/forms';


@NgModule({
  declarations: [
    GraficoBarrasEvolutivoMarcasComponent
  ],
  imports: [
    CommonModule,
    SelectImageModule,
    NgSelectModule,
    FormsModule ,
  ],
  exports: [GraficoBarrasEvolutivoMarcasComponent]
})

export class GraficoBarrasEvolutivoMarcasModule { }
