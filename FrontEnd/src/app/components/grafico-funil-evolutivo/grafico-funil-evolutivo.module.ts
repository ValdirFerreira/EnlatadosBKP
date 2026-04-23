import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { NgSelectModule } from '@ng-select/ng-select';
import { GraficoFunilEvolutivoComponent } from './grafico-funil-evolutivo.component';
import { SelectImageModule } from '../select-image/select-image.module';
import { TranslateModule } from '@ngx-translate/core';
import { FormsModule } from '@angular/forms';

@NgModule({
  declarations: [
    GraficoFunilEvolutivoComponent
  ],
  imports: [
    CommonModule,
    NgSelectModule,
    SelectImageModule,
    TranslateModule,
    FormsModule ,
  ],
  exports: [GraficoFunilEvolutivoComponent]
})

export class GraficoFunilEvolutivoModule { }
