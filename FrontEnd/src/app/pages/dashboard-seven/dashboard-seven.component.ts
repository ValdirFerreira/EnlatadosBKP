import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import * as Highcharts from 'highcharts';
import { FiltroPadrao, FiltroPadraoExcel } from 'src/app/models/Filtros/FiltroPadrao';
import { saveAs } from "file-saver";
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';
import { MenuService } from 'src/app/services/menu.service';
import { DownloadArquivoService } from 'src/app/services/download-arquivo.service';
import { EventEmitterService } from 'src/app/services/event-emitter.service';
import { GraficoColunaModel } from 'src/app/models/grafico-coluna/grafico-coluna';
import { DashBoardTwoService } from 'src/app/services/dashboard-two.service';
import { TranslateService } from '@ngx-translate/core';
import { ConversorPowerpointService } from 'src/app/services/conversor-powerpoint.service';
import { AuthService } from 'src/app/services/auth.service';
import { Session } from '../home/guards/session';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';
import { DashBoardSevenService } from 'src/app/services/dashboard-seven.service';
import { GraficoComunicacaoVisto } from 'src/app/models/GraficoComunicacaoVisto/GraficoComunicacaoVisto';
import { GraficoComunicacaoRecall } from 'src/app/models/GraficoComunicacaoRecall/GraficoComunicacaoRecall';
import { GraficoComunicacaoDiagnostico } from 'src/app/models/GraficoComunicacaoDiagnostico/GraficoComunicacaoDiagnostico';
import { ComunicacaoQuadroResumo } from 'src/app/models/ComunicacaoQuadroResumo/ComunicacaoQuadroResumo';
import { LogService } from 'src/app/services/log.service';



@Component({
  selector: 'dashboard-seven',
  templateUrl: './dashboard-seven.component.html',
  styleUrls: ['./dashboard-seven.component.scss']
})
export class DashboardSevenComponent implements OnInit {

  graficoComunicacaoVisto = new Array<GraficoComunicacaoVisto>();
  graficoComunicacaoRecall = new Array<GraficoComunicacaoRecall>();
  graficoComunicacaoDiagnostico = new Array<GraficoComunicacaoDiagnostico>();
  graficoComunicacaoQuadroResumo = new Array<ComunicacaoQuadroResumo>();

  graficoComunicacaoSource = new Array<GraficoComunicacaoVisto>();

  ativaGraficoComparativoMarcas: boolean = false;
  ativaGraficoEvolutivoMarcas: boolean = false;

  ativaGraficoImagens: boolean = false;
  ModelSTB: PadraoComboFiltro;
  idItemSTBFoto: number = 0;
  public listaAtributoSTB: Array<PadraoComboFiltro>;

  constructor(public router: Router,
    public menuService: MenuService,
    public filtroService: FiltroGlobalService, private downloadArquivoService: DownloadArquivoService, private dashBoardSevenService: DashBoardSevenService,
    private translate: TranslateService, private conversorPowerpointService: ConversorPowerpointService,
    private authService: AuthService,
    private session: Session,
    private logService: LogService,
  ) { }

  ngOnInit(): void {

    window.scroll({
      top: 0,
      left: 0,
      behavior: 'smooth'
    });

    this.menuService.nomePage = this.translate.instant('navbar.dashboard-seven');
    this.menuService.activeMenu = 7;
    this.menuService.menuSelecao = "7"

    EventEmitterService.get("emit-dashboard-seven").subscribe((x) => {
      this.carregaFiltroSTB();

      this.logService.GravaLogRota(this.router.url).subscribe(
      );

    })

  }

  carregaFiltroSTB() {
    this.filtroService.FiltroSTB(this.filtroService.ModelOnda)
      .subscribe((response: Array<PadraoComboFiltro>) => {
        this.listaAtributoSTB = response;
        if (this.listaAtributoSTB && this.listaAtributoSTB.length > 0)
          this.ModelSTB = this.listaAtributoSTB[0];
        else {
          this.ModelSTB = null;
        }

        if (this.ModelSTB)
          this.idItemSTBFoto = this.ModelSTB.IdItem;
        else {
          this.idItemSTBFoto = 0;
        }

        this.carregaImgens();

      }, (error) => console.error(error),
        () => {
        }
      )
  }

  public carregaFiltros() {


    var filtros = new FiltroPadrao();
    var list =[];
    list.push(this.filtroService.ModelTarget);
    filtros.Target = list;
    filtros.Regiao = this.filtroService.ModelRegiao;
    filtros.Demografico = this.filtroService.ModelDemografico;
    filtros.Onda.push(this.filtroService.ModelOnda);
    filtros.CodUser = this.session.getCodUserSession();
    filtros.CodIdioma = 1;// this.authService.idDefaultLangUser;
    filtros.ParamSTB = this.ModelSTB ? this.ModelSTB.IdItem : 0;
    return filtros;
  }

  carregaImgens() {
    this.ativaGraficoImagens = true;
    if (this.ModelSTB) {
      var filtro = this.carregaFiltros();
      filtro.Sequencia = 1;
      filtro.ParamSTB = this.ModelSTB.IdItem;
      this.dashBoardSevenService.carregarComunicacaoQuadroResumo(filtro)
        .subscribe((grafico: Array<ComunicacaoQuadroResumo>) => {
          this.graficoComunicacaoQuadroResumo = grafico;

          this.carregarGraficoComunicacaoVisto();
          this.carregarGraficoComunicacaoDiagnostico();
          this.carregarGraficoComunicacaoRecall();

           this.carregarGraficoComunicacaoSource();

        }, (error) => { console.error(error) },
          () => { })
    }
    else {
      this.carregarGraficoComunicacaoVisto();
      this.carregarGraficoComunicacaoDiagnostico();
      this.carregarGraficoComunicacaoRecall();

       this.carregarGraficoComunicacaoSource();
    }
  }

  carregarGraficoComunicacaoRecall() {

    var filtro = this.carregaFiltros();
    filtro.Sequencia = 1;
    this.dashBoardSevenService.carregarGraficoComunicacaoRecall(filtro)
      .subscribe((grafico: Array<GraficoComunicacaoRecall>) => {
        this.graficoComunicacaoRecall = grafico
      }, (error) => { console.error(error) },
        () => { })
  }

  carregarGraficoComunicacaoVisto() {

    var filtro = this.carregaFiltros();
    filtro.Sequencia = 1;
    this.dashBoardSevenService.carregarGraficoComunicacaoVisto(filtro)
      .subscribe((grafico: Array<GraficoComunicacaoVisto>) => {
        this.graficoComunicacaoVisto = grafico;

      }, (error) => { console.error(error) },
        () => { })
  }

  carregarGraficoComunicacaoDiagnostico() {

    var filtro = this.carregaFiltros();
    filtro.Sequencia = 1;
    this.dashBoardSevenService.carregarGraficoComunicacaoDiagnostico(filtro)
      .subscribe((grafico: Array<GraficoComunicacaoDiagnostico>) => {
        this.graficoComunicacaoDiagnostico = grafico
      }, (error) => { console.error(error) },
        () => { })
  }

  onchangeSTB(ModelSTB: PadraoComboFiltro) {
    this.idItemSTBFoto = this.ModelSTB.IdItem;
    this.carregaImgens();
    this.carregarGraficoComunicacaoRecall();
    this.carregarGraficoComunicacaoVisto();
    this.carregarGraficoComunicacaoDiagnostico();
    this.carregarGraficoComunicacaoSource();
  }


  carregarGraficoComunicacaoSource() {

    var filtro = this.carregaFiltros();
    filtro.Sequencia = 10;
    this.dashBoardSevenService.carregarGraficoComunicacaoSource(filtro)
      .subscribe((grafico: Array<GraficoComunicacaoVisto>) => {
        this.graficoComunicacaoSource = grafico;

      }, (error) => { console.error(error) },
        () => { })
  }

  ajustaScrollFinal() {
    if (document.getElementById('aqui_id_div_grafico') != null && !document.getElementById('aqui_id_div_grafico').scrollLeft) {
      document.getElementById('aqui_id_div_grafico').scrollLeft = 9999;
    }

  }

  exportToPpt1() {

    var nome = this.translate.instant('dashboard-seven-grafico1-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'grafico1',
      nome,
    );
  }

  exportToPpt3() {
    var nome = this.translate.instant('dashboard-seven-grafico3-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'grafico3',
      nome
    );
  }




  exportToPpt4() {
    var nome = this.translate.instant('dashboard-seven-grafico4-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'grafico4',
      nome
    );
  }


  exportToPpt5() {
    var nome = this.translate.instant('dashboard-seven-grafico5-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'grafico5',
      nome
    );
  }

  async exportToExcelGrafico1() {

    var filtroPadrao = this.carregaFiltros();
    var titulo = this.translate.instant('dashboard-seven-grafico1-titulo');
    filtroPadrao.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardGraficoComunicacaoRecall(filtroPadrao)
      .subscribe((result) => {

        saveAs(
          result,
          titulo + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => console.error(error));

  }

  async exportToExcelGrafico2() {
    var filtroPadrao = this.carregaFiltros();
    var titulo = this.translate.instant('dashboard-seven-grafico3-titulo');
    filtroPadrao.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardGraficoComunicacaoVisto(filtroPadrao)
      .subscribe((result) => {

        saveAs(
          result,
          titulo + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => console.error(error));
  }

  async exportToExcelGrafico3() {
    var filtroPadrao = this.carregaFiltros();
    var titulo = this.translate.instant('dashboard-seven-grafico4-titulo');
    filtroPadrao.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardGraficoComunicacaoDiagnostico(filtroPadrao)
      .subscribe((result) => {

        saveAs(
          result,
          titulo + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => console.error(error));
  }


  async exportToExcelGrafico4() {
    var filtroPadrao = this.carregaFiltros();
    var titulo = this.translate.instant('dashboard-seven-grafico5-titulo');
    filtroPadrao.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardGraficoComunicacaoSource(filtroPadrao)
      .subscribe((result) => {

        saveAs(
          result,
          titulo + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => console.error(error));
  }



  
}
