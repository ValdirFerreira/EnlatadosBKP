import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { SelectImageModule } from '../select-image/select-image.module';
import { TranslateModule } from '@ngx-translate/core';
import { GraficoColunasSimplesComponent } from './grafico-colunas-simples.component';


@NgModule({
  declarations: [
    GraficoColunasSimplesComponent
  ],
  imports: [
    CommonModule,
    SelectImageModule,
    TranslateModule,
  ],
  exports: [GraficoColunasSimplesComponent]
})

export class GraficoColunasSimplesModule { }
