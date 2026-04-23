import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { GraficoFunilComponent } from './grafico-funil.component';
import { NgSelectModule } from '@ng-select/ng-select';
import { SelectImageModule } from '../select-image/select-image.module';
import { TranslateModule } from '@ngx-translate/core';

@NgModule({
  declarations: [
    GraficoFunilComponent
  ],
  imports: [
    CommonModule,
    NgSelectModule,
    SelectImageModule,
    TranslateModule
  ],
  exports: [GraficoFunilComponent]
})

export class GraficoFunilModule { }
