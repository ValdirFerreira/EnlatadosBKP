import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { GraficoColunasMarcas4RowComponent } from './grafico-colunas-marcas-4-row.component';
import { NgSelectModule } from '@ng-select/ng-select';
import { SelectImageModule } from '../select-image/select-image.module';
import { FormsModule } from '@angular/forms';


@NgModule({
  declarations: [
    GraficoColunasMarcas4RowComponent
  ],
  imports: [
    CommonModule,
    NgSelectModule,
    SelectImageModule,
    FormsModule ,
  ],
  exports: [GraficoColunasMarcas4RowComponent]
})

export class GraficoColunasMarcas4RowModule { }
