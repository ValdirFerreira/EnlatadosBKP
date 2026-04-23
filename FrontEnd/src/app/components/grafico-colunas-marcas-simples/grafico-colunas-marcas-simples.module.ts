import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { GraficoColunasMarcasSimplesComponent } from './grafico-colunas-marcas-simples.component';
import { NgSelectModule } from '@ng-select/ng-select';
import { SelectImageModule } from '../select-image/select-image.module';
import { FormsModule } from '@angular/forms';


@NgModule({
  declarations: [
    GraficoColunasMarcasSimplesComponent
  ],
  imports: [
    CommonModule,
    NgSelectModule,
    SelectImageModule,
    FormsModule ,
  ],
  exports: [GraficoColunasMarcasSimplesComponent]
})

export class GraficoColunasMarcasSimplesModule { }
