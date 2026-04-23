import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { SelectImageModule } from '../select-image/select-image.module';
import { GraficoBarrasTopEvolutivoComponent } from './grafico-barras-top-evolutivo.component';
import { TranslateModule } from '@ngx-translate/core';
import { NgSelectModule } from '@ng-select/ng-select';
import { FormsModule } from '@angular/forms';



@NgModule({
  declarations: [
    GraficoBarrasTopEvolutivoComponent
  ],
  imports: [
    CommonModule,
    SelectImageModule,
    TranslateModule,
     NgSelectModule,
    FormsModule ,
  ],
  exports: [GraficoBarrasTopEvolutivoComponent]
})

export class GraficoBarrasTopEvolutivoModule { }
