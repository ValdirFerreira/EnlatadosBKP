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
import { GraficoFunilModule } from 'src/app/components/grafico-funil/grafico-funil.module';
import { GraficoFunilEvolutivoModule } from 'src/app/components/grafico-funil-evolutivo/grafico-funil-evolutivo.module';
import { TranslateModule } from '@ngx-translate/core';
import { DashboardNineComponent } from './dashboard-nine.component';
import { DashboardNineRoutingModule } from './dashboard-nine-routing.module';

import { MatCheckboxModule } from '@angular/material/checkbox';
import { FormsModule } from '@angular/forms';
import { NgSelectModule } from '@ng-select/ng-select';
import { SelectImageModule } from 'src/app/components/select-image/select-image.module';
import { SelectCheckboxModule } from 'src/app/components/select-checkbox/select-checkbox.module';
import { SelectCheckboxAtributoModule } from 'src/app/components/select-checkbox-atributo/select-checkbox-atributo.module';
import { MatRadioModule } from '@angular/material/radio';
import { SelectImage2Module } from 'src/app/components/select-image-2/select-image-2.module';
import { TabelaPadraoModule } from 'src/app/components/tabela-padrao/tabela-padrao.module';
import { SelectCheckboxAdHocModule } from 'src/app/components/select-checkbox-adHoc/select-checkbox-adHoc.module';
import { TabelaPadraoMarcaModule } from 'src/app/components/tabela-padrao-base-marca/tabela-padrao-marca.module';
import { TabelaPadraoBaseModule } from 'src/app/components/tabela-padrao-base/tabela-padrao-base.module';

@NgModule({
  providers: [
    { provide: HIGHCHARTS_MODULES, useFactory: () => [more, exporting, solidgauge, wordcloud, treemap] } // add as factory to your providers
  ],
  declarations: [
    DashboardNineComponent,
  ],
  imports: [
    CommonModule,
    DashboardNineRoutingModule,
    SidebarModule,
    NavbarModule,
    AvisoSemDadosModule,
    ChartModule,
    FooterBottomModule,
    GraficoFunilModule,
    GraficoFunilEvolutivoModule,
    TranslateModule,
    MatCheckboxModule,
    FormsModule,
    NgSelectModule,
    SelectImageModule,
    SelectCheckboxModule,
    SelectCheckboxAtributoModule,
    MatRadioModule,
    SelectImage2Module,
    TabelaPadraoModule,
    SelectCheckboxAdHocModule,
    TabelaPadraoMarcaModule,
    TabelaPadraoBaseModule
  ]
})
export class DashboardNineModule { }
