import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { GraficoDiagnosticoPropagandaComponent } from './grafico-diagnostico-propaganda.component';
import { TranslateModule } from '@ngx-translate/core';

@NgModule({
  declarations: [
    GraficoDiagnosticoPropagandaComponent
  ],
  imports: [
    CommonModule,
    TranslateModule
  ],
  exports: [GraficoDiagnosticoPropagandaComponent]
})

export class GraficoDiagnosticoPropagandaModule { }
