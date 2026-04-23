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
import {  DashboardSixRoutingModule } from './dashboard-six-routing.module';
import { GraficoBarrasMarcasModule } from 'src/app/components/grafico-barras-marcas/grafico-barras-marcas.module';
import { GraficoBarrasEvolutivoMarcasModule } from 'src/app/components/grafico-barras-evolutivo-marcas/grafico-barras-evolutivo-marcas.module';
import { GraficoBarrasMarcasDuploModule } from 'src/app/components/grafico-barras-marcas-duplo/grafico-barras-marcas-duplo.module';
import { DashboardSixComponent } from './dashboard-six.component';
import { GraficoBarrasTopMarcasModule } from 'src/app/components/grafico-barras-top-marcas/grafico-barras-top-marcas.module';
import { GraficoBarrasTopEvolutivoModule } from 'src/app/components/grafico-barras-top-evolutivo/grafico-barras-top-evolutivo.module';
import { GraficoBarrasMarcasDuploMercadoModule } from 'src/app/components/grafico-barras-marcas-duplo-mercado/grafico-barras-marcas-duplo-mercado.module';
import { GraficoBarrasEvolutivoMarcasMercadoModule } from 'src/app/components/grafico-barras-evolutivo-marcas-mercado/grafico-barras-evolutivo-marcas-mercado.module';
import { SelectCheckboxBVCModule } from 'src/app/components/select-checkbox-bvc/select-checkbox-bvc.module';
import { NgSelectModule } from '@ng-select/ng-select';
import { FormsModule } from '@angular/forms';


@NgModule({
  providers: [
    { provide: HIGHCHARTS_MODULES, useFactory: () => [more, exporting, solidgauge, wordcloud, treemap] } // add as factory to your providers
  ],
  declarations: [
    DashboardSixComponent

  ],
  imports: [
    CommonModule,
    DashboardSixRoutingModule,
    SidebarModule,
    NavbarModule,
    AvisoSemDadosModule,
    ChartModule,
    FooterBottomModule,
    TranslateModule,
    SelectImageModule,
    GraficoBarrasTopMarcasModule,
    GraficoBarrasTopEvolutivoModule,
    GraficoBarrasMarcasDuploMercadoModule,
    GraficoBarrasEvolutivoMarcasMercadoModule,
    SelectCheckboxBVCModule,
    NgSelectModule,
    FormsModule ,
  ]
})
export class DashboardSixModule { }
