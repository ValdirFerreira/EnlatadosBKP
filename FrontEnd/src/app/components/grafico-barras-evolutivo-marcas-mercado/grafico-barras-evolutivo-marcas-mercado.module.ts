import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { SelectImageModule } from '../select-image/select-image.module';
import { GraficoBarrasEvolutivoMarcasMercadoComponent } from './grafico-barras-evolutivo-marcas-mercado.component';
import { NgSelectModule } from '@ng-select/ng-select';
import { FormsModule } from '@angular/forms';


@NgModule({
  declarations: [
    GraficoBarrasEvolutivoMarcasMercadoComponent
  ],
  imports: [
    CommonModule,
    SelectImageModule,
    NgSelectModule,
    FormsModule ,
  ],
  exports: [GraficoBarrasEvolutivoMarcasMercadoComponent]
})

export class GraficoBarrasEvolutivoMarcasMercadoModule { }
