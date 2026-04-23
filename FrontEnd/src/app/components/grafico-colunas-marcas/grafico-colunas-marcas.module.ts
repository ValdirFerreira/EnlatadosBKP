import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { GraficoColunasMarcasComponent } from './grafico-colunas-marcas.component';
import { NgSelectModule } from '@ng-select/ng-select';
import { SelectImageModule } from '../select-image/select-image.module';
import { FormsModule } from '@angular/forms';


@NgModule({
  declarations: [
    GraficoColunasMarcasComponent
  ],
  imports: [
    CommonModule,
    NgSelectModule,
    SelectImageModule,
    FormsModule ,
  ],
  exports: [GraficoColunasMarcasComponent]
})

export class GraficoColunasMarcasModule { }
