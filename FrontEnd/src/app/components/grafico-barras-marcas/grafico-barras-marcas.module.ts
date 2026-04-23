import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { GraficoBarrasMarcasComponent } from './grafico-barras-marcas.component';
import { SelectImageModule } from '../select-image/select-image.module';


@NgModule({
  declarations: [
    GraficoBarrasMarcasComponent
  ],
  imports: [
    CommonModule,
    SelectImageModule
  ],
  exports: [GraficoBarrasMarcasComponent]
})

export class GraficoBarrasMarcasModule { }
