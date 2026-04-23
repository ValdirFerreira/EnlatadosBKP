import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { SelectImageModule } from '../select-image/select-image.module';
import { GraficoBarrasTopMarcasComponent } from './grafico-barras-top-marcas.component';
import { TranslateModule } from '@ngx-translate/core';


@NgModule({
  declarations: [
    GraficoBarrasTopMarcasComponent
  ],
  imports: [
    CommonModule,
    SelectImageModule,
    TranslateModule,
  ],
  exports: [GraficoBarrasTopMarcasComponent]
})

export class GraficoBarrasTopMarcasModule { }
