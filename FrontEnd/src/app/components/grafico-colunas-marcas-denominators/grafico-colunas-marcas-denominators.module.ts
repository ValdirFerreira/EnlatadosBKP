import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { GraficoColunasMarcasDenominatorsComponent } from './grafico-colunas-marcas-denominators.component';
import { NgSelectModule } from '@ng-select/ng-select';
import { SelectImageModule } from '../select-image/select-image.module';
import { FormsModule } from '@angular/forms';


@NgModule({
  declarations: [
    GraficoColunasMarcasDenominatorsComponent
  ],
  imports: [
    CommonModule,
    NgSelectModule,
    SelectImageModule,
    FormsModule ,
  ],
  exports: [GraficoColunasMarcasDenominatorsComponent]
})

export class GraficoColunasMarcasDenominatorsModule { }
