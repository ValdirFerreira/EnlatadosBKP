import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { SidebarModule } from 'src/app/components/sidebar/sidebar.module';
import { NavbarModule } from 'src/app/components/navbar/navbar.module';
import { FiltroGlobalModule } from 'src/app/components/filtroGlobal/filtro-global.module';
import { AvisoSemDadosModule } from 'src/app/components/aviso-sem-dados/aviso-sem-dados.module';

import { ChartModule, HIGHCHARTS_MODULES } from 'angular-highcharts';
import * as more from 'highcharts/highcharts-more.src';
import * as exporting from 'highcharts/modules/exporting.src';
import * as solidgauge from 'highcharts/modules/solid-gauge.src';
import * as wordcloud from 'highcharts/modules/wordcloud.src';
import * as treemap from 'highcharts/modules/treemap.src';
import * as data from 'highcharts/modules/data.src';

import { FooterBottomModule } from 'src/app/components/footer-bottom/footer-bottom.module';
import { DashboardSevenRoutingModule } from './dashboard-seven-routing.module';
import { DashboardSevenComponent } from './dashboard-seven.component';
import { GraficoColunasModule } from 'src/app/components/grafico-colunas/grafico-colunas.module';
import { GraficoColunasMarcasModule } from 'src/app/components/grafico-colunas-marcas/grafico-colunas-marcas.module';
import { GraficoColunasEvolutivoMarcasModule } from 'src/app/components/grafico-colunas-evolutivo-marcas/grafico-colunas-evolutivo-marcas.module';
import { TranslateModule } from '@ngx-translate/core';
import { SelectImageModule } from 'src/app/components/select-image/select-image.module';
import { GraficoColunaDuploModule } from 'src/app/components/grafico-coluna-duplo/grafico-coluna-duplo.module';
import { GraficoPropagandaModule } from 'src/app/components/grafico-propaganda/grafico-propaganda.module';
import { GraficoColunasSimplesModule } from 'src/app/components/grafico-colunas-simples/grafico-colunas-simples.module';
import { GraficoDiagnosticoPropagandaModule } from 'src/app/components/grafico-diagnostico-propaganda/grafico-diagnostico-propaganda.module';
import { SelectCheckboxSTBModule } from 'src/app/components/select-checkbox-STB/select-checkbox-STB.module';
import { NgSelectModule } from '@ng-select/ng-select';
import { FormsModule } from '@angular/forms';


@NgModule({
  providers: [
    { provide: HIGHCHARTS_MODULES, useFactory: () => [more, exporting, solidgauge, wordcloud, treemap] } // add as factory to your providers
  ],
  declarations: [
    DashboardSevenComponent,

  ],
  imports: [
    CommonModule,
    DashboardSevenRoutingModule,
    SidebarModule,
    NavbarModule,
    AvisoSemDadosModule,
    ChartModule,
    FooterBottomModule,
    GraficoColunasModule,
    GraficoColunasMarcasModule,
    GraficoColunasEvolutivoMarcasModule,
    TranslateModule,
    SelectImageModule,
    GraficoColunaDuploModule,
    GraficoPropagandaModule,
    GraficoColunasSimplesModule,
    GraficoDiagnosticoPropagandaModule,
    SelectCheckboxSTBModule,
    NgSelectModule,
    FormsModule ,
  ]
})
export class DashboardSevenModule { }
