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
import { TranslateModule } from '@ngx-translate/core';
import { SelectImageModule } from 'src/app/components/select-image/select-image.module';
import { DashboardFiveComponent } from './dashboard-five.component';
import { DashboardFiveRoutingModule } from './dashboard-five-routing.module';
import { GraficoBarrasMarcasModule } from 'src/app/components/grafico-barras-marcas/grafico-barras-marcas.module';
import { GraficoBarrasEvolutivoMarcasModule } from 'src/app/components/grafico-barras-evolutivo-marcas/grafico-barras-evolutivo-marcas.module';
import { GraficoBarrasMarcasDuploModule } from 'src/app/components/grafico-barras-marcas-duplo/grafico-barras-marcas-duplo.module';
import { GraficoBarrasMarcasTriploModule } from 'src/app/components/grafico-barras-marcas-triplo/grafico-barras-marcas-triplo.module';


@NgModule({
  providers: [
    { provide: HIGHCHARTS_MODULES, useFactory: () => [more, exporting, solidgauge, wordcloud, treemap] } // add as factory to your providers
  ],
  declarations: [
    DashboardFiveComponent,

  ],
  imports: [
    CommonModule,
    DashboardFiveRoutingModule,
    SidebarModule,
    NavbarModule,
    AvisoSemDadosModule,
    ChartModule,
    FooterBottomModule,
    TranslateModule,
    SelectImageModule,
    GraficoBarrasMarcasModule,
    GraficoBarrasEvolutivoMarcasModule,
    GraficoBarrasMarcasDuploModule,
    GraficoBarrasMarcasTriploModule
  ]
})
export class DashboardFiveModule { }
