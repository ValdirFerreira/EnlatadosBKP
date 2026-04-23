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
import { stringify } from 'querystring';
import { LogService } from 'src/app/services/log.service';



@Component({
  selector: 'app-dashboard-two',
  templateUrl: './dashboard-two.component.html',
  styleUrls: ['./dashboard-two.component.scss']
})
export class DashboardTwoComponent implements OnInit {

  graficoColunaModel = new Array<GraficoColunaModel>();

  graficoColunaModel1 = new GraficoColunaModel();
  graficoColunaModel2 = new GraficoColunaModel();
  graficoColunaModel3 = new GraficoColunaModel();
  graficoColunaModel4 = new GraficoColunaModel();
  graficoColunaModel5 = new GraficoColunaModel();

  grafico2ColunaModel1 = new GraficoColunaModel();
  grafico2ColunaModel2 = new GraficoColunaModel();
  grafico2ColunaModel3 = new GraficoColunaModel();
  grafico2ColunaModel4 = new GraficoColunaModel();
  grafico2ColunaModel5 = new GraficoColunaModel();

  // Filtros de Marca para utilização da geração do Excel Grafico Comparativo Marcas 
  marcaColuna1: PadraoComboFiltro;
  marcaColuna2: PadraoComboFiltro;
  marcaColuna3: PadraoComboFiltro;
  marcaColuna4: PadraoComboFiltro;
  marcaColuna5: PadraoComboFiltro;
  // Filtros de Marca para utilização da geração do Excel Grafico Comparativo Marcas 

  // Filtros de onde para utilização da geração do Excel Grafico Marcas Evolutivo
  ondaColuna1: PadraoComboFiltro;
  ondaColuna2: PadraoComboFiltro;
  ondaColuna3: PadraoComboFiltro;
  ondaColuna4: PadraoComboFiltro;
  marcaEvolutivo: PadraoComboFiltro;
  // Filtros de onde para utilização da geração do Excel Grafico Marcas Evolutivo

  subscription: any;

  ativaGraficoComparativoMarcas: boolean = false;
  ativaGraficoComparativoMarcasDenominators: boolean = false;
  ativaGraficoEvolutivoMarcas: boolean = false;
  ativaGraficoEvolutivoMarcasDenominators: boolean = false;

  constructor(public router: Router,
    public menuService: MenuService,
    public filtroService: FiltroGlobalService, private downloadArquivoService: DownloadArquivoService, private dashBoardTwoService: DashBoardTwoService,
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
    this.menuService.nomePage = this.translate.instant('navbar.dashboard-two');
    this.menuService.activeMenu = 2;
    this.menuService.menuSelecao = "2"

    this.subscription = EventEmitterService.get("emit-dashboard-two").subscribe((x) => {

      // this.ativaGraficoComparativoMarcas = false;
      // this.ativaGraficoEvolutivoMarcas = false;  


      this.carregarGraficos();

      this.logService.GravaLogRota(this.router.url).subscribe(
      );
    })


  }

  ngOnDestroy() {
    this.subscription.unsubscribe();
  }


  carregarGraficos() {

    this.filtroService.FiltroMarcas(this.filtroService.ModelRegiao)
      .subscribe((response: Array<PadraoComboFiltro>) => {
        this.filtroService.listaMarcas = response;

        this.marcaColuna1 = null;
        this.marcaColuna2 = null;
        this.marcaColuna3 = null;
        this.marcaColuna4 = null;
        this.marcaColuna5 = null;

        this.ondaColuna1 = null;
        this.ondaColuna2 = null;
        this.ondaColuna3 = null;
        this.ondaColuna4 = null;

        this.ativaGraficoComparativoMarcas = false;
        this.ativaGraficoEvolutivoMarcas = false;
        this.ativaGraficoEvolutivoMarcasDenominators = false;
        this.ativaGraficoComparativoMarcasDenominators = false;

        this.grafico2ColunaModel1 = null;
        this.grafico2ColunaModel2 = null;
        this.grafico2ColunaModel3 = null;
        this.grafico2ColunaModel4 = null;

        this.carregaFiltroMarcas();
        this.carregaFiltroOndas();
        this.carregarGraficoComparativoMarcasFirstLoad();
        this.carregarGraficoEvolutivoMarcasFirstLoad(null);
        // this.carregarGraficoEvolutivoMarcasFirstLoad(this.marcaEvolutivo);


      }, (error) => console.error(error),
        () => {
        }
      )
  }


  public carregaFiltros() {

    var filtros = new FiltroPadrao();
    var list = [];
    list.push(this.filtroService.ModelTarget);
    filtros.Target = list;
    filtros.Regiao = this.filtroService.ModelRegiao;
    filtros.Demografico = this.filtroService.ModelDemografico;
    filtros.Onda.push(this.filtroService.ModelOnda);
    filtros.CodUser = this.session.getCodUserSession();
    filtros.CodIdioma = 1;// this.authService.idDefaultLangUser;
    filtros.ParamDenominators = 1;// this.filtroService.ModelTarget.IdItem;
    return filtros;
  }


  public carregaFiltrosExcel() {

    var filtros = new FiltroPadraoExcel();
    var list = [];
    list.push(this.filtroService.ModelTarget);
    filtros.Target = list;
    filtros.Regiao = this.filtroService.ModelRegiao;
    filtros.Demografico = this.filtroService.ModelDemografico;
    filtros.Onda.push(this.filtroService.ModelOnda);

    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Marca.push(this.marcaEvolutivo);


    filtros.Onda1.push(this.ondaColuna1);
    filtros.Onda2.push(this.ondaColuna2);
    filtros.Onda3.push(this.ondaColuna3);
    filtros.Onda4.push(this.ondaColuna4);

    filtros.Marca1.push(this.marcaColuna1);
    filtros.Marca2.push(this.marcaColuna2);
    filtros.Marca3.push(this.marcaColuna3);
    filtros.Marca4.push(this.marcaColuna4);
    filtros.Marca5.push(this.marcaColuna5);

    filtros.ParamDenominators = 1;// this.filtroService.ModelTarget.IdItem;

    filtros.CodUser = this.session.getCodUserSession();
    filtros.CodIdioma = 1;// this.authService.idDefaultLangUser;
    return filtros;
  }

  public carregaFiltroMarcas() {

    if (!this.marcaColuna1)
      this.marcaColuna1 = this.filtroService.listaMarcas[0]

    if (!this.marcaColuna2)
      this.marcaColuna2 = this.filtroService.listaMarcas[1]

    if (!this.marcaColuna3)
      this.marcaColuna3 = this.filtroService.listaMarcas[2]

    if (!this.marcaColuna4)
      this.marcaColuna4 = this.filtroService.listaMarcas[3]

    if (!this.marcaColuna5)
      this.marcaColuna5 = this.filtroService.listaMarcas[4]


  }

  public carregaFiltroOndas() {
    if (!this.ondaColuna1) {
      this.ondaColuna1 = this.filtroService.listaOnda[3];
    }

    if (!this.ondaColuna2) {
      this.ondaColuna2 = this.filtroService.listaOnda[2];
    }

    if (!this.ondaColuna3) {
      this.ondaColuna3 = this.filtroService.listaOnda[1];
    }

    if (!this.ondaColuna4) {
      this.ondaColuna4 = this.filtroService.listaOnda[0];
    }

  }

  carregarGraficoComparativoMarcasFirstLoad() {




    var filtroColuna1 = this.carregaFiltros();
    filtroColuna1.Marca = new Array<PadraoComboFiltro>();
    filtroColuna1.Marca.push(this.marcaColuna1);
    filtroColuna1.Sequencia = 1;
    filtroColuna1.ParamDenominators = 1;// this.filtroService.ModelTarget.IdItem;
    this.dashBoardTwoService.carregarGraficoComparativoMarcas(filtroColuna1)
      .subscribe((coluna1: GraficoColunaModel) => {
        this.graficoColunaModel1 = coluna1
      }, (error) => console.error(error),
        () => { })

    var filtroColuna2 = this.carregaFiltros();
    filtroColuna2.Sequencia = 2;
    filtroColuna2.Marca = new Array<PadraoComboFiltro>();
    filtroColuna2.Marca.push(this.marcaColuna2);
    filtroColuna2.ParamDenominators = 1;// this.filtroService.ModelTarget.IdItem;
    this.dashBoardTwoService.carregarGraficoComparativoMarcas(filtroColuna2)
      .subscribe((coluna2: GraficoColunaModel) => {
        this.graficoColunaModel2 = coluna2

      }, (error) => console.error(error),
        () => { })

    var filtroColuna3 = this.carregaFiltros();
    filtroColuna3.Sequencia = 3;
    filtroColuna3.Marca = new Array<PadraoComboFiltro>();
    filtroColuna3.Marca.push(this.marcaColuna3);
    filtroColuna3.ParamDenominators = 1;// this.filtroService.ModelTarget.IdItem;
    this.dashBoardTwoService.carregarGraficoComparativoMarcas(filtroColuna3)
      .subscribe((coluna3: GraficoColunaModel) => {
        this.graficoColunaModel3 = coluna3

      }, (error) => console.error(error),
        () => { })

    var filtroColuna4 = this.carregaFiltros();
    filtroColuna4.Sequencia = 4;
    filtroColuna4.Marca = new Array<PadraoComboFiltro>();
    filtroColuna4.Marca.push(this.marcaColuna4);
    filtroColuna4.ParamDenominators =1;// this.filtroService.ModelTarget.IdItem;
    this.dashBoardTwoService.carregarGraficoComparativoMarcas(filtroColuna4)
      .subscribe((coluna4: GraficoColunaModel) => {
        this.graficoColunaModel4 = coluna4

      }, (error) => console.error(error),
        () => { })

    var filtroColuna5 = this.carregaFiltros();
    filtroColuna5.Sequencia = 5;
    filtroColuna5.Marca = new Array<PadraoComboFiltro>();
    filtroColuna5.Marca.push(this.marcaColuna5);
    filtroColuna5.ParamDenominators = 1;// this.filtroService.ModelTarget.IdItem;
    this.dashBoardTwoService.carregarGraficoComparativoMarcas(filtroColuna5)
      .subscribe((coluna5: GraficoColunaModel) => {
        this.graficoColunaModel5 = coluna5

            this.ativaGraficoComparativoMarcas = true;
   
        // if (this.filtroService.ModelTarget.IdItem == 1) {
        //   this.ativaGraficoComparativoMarcas = true;
        // }
        // if (this.filtroService.ModelTarget.IdItem == 2) {
        //   this.ativaGraficoComparativoMarcasDenominators = true;
        // }

      }, (error) => console.error(error),
        () => { })

  }

  carregarGraficoEvolutivoMarcasFirstLoad(item: PadraoComboFiltro = null) {


    var marca = item == null ? this.filtroService.listaMarcas[0] : item;

    var filtroColuna1 = this.carregaFiltros();
    filtroColuna1.Sequencia = 6;
    filtroColuna1.Marca = new Array<PadraoComboFiltro>();
    filtroColuna1.Marca.push(marca);
    filtroColuna1.Onda = new Array<PadraoComboFiltro>();
    filtroColuna1.Onda.push(this.ondaColuna1);
    this.marcaEvolutivo = marca;
    filtroColuna1.ParamDenominators =1;// this.filtroService.ModelTarget.IdItem;

    this.dashBoardTwoService.carregarGraficoComparativoMarcas(filtroColuna1)
      .subscribe((coluna1: GraficoColunaModel) => {
        this.grafico2ColunaModel1 = coluna1
      }, (error) => console.error(error),
        () => { })

    var filtroColuna2 = this.carregaFiltros();
    filtroColuna2.Sequencia = 7;
    filtroColuna2.Marca = new Array<PadraoComboFiltro>();
    filtroColuna2.Marca.push(marca);
    filtroColuna2.Onda = new Array<PadraoComboFiltro>();
    filtroColuna2.Onda.push(this.ondaColuna2);
    filtroColuna2.ParamDenominators  =1;// this.filtroService.ModelTarget.IdItem;
    this.dashBoardTwoService.carregarGraficoComparativoMarcas(filtroColuna2)
      .subscribe((coluna2: GraficoColunaModel) => {
        this.grafico2ColunaModel2 = coluna2
      }, (error) => console.error(error),
        () => { })

    var filtroColuna3 = this.carregaFiltros();
    filtroColuna3.Sequencia = 8;
    filtroColuna3.Marca = new Array<PadraoComboFiltro>();
    filtroColuna3.Marca.push(marca);
    filtroColuna3.Onda = new Array<PadraoComboFiltro>();
    filtroColuna3.Onda.push(this.ondaColuna3);
    filtroColuna3.ParamDenominators = 1;// this.filtroService.ModelTarget.IdItem;
    this.dashBoardTwoService.carregarGraficoComparativoMarcas(filtroColuna3)
      .subscribe((coluna3: GraficoColunaModel) => {
        this.grafico2ColunaModel3 = coluna3

      }, (error) => console.error(error),
        () => { })

    var filtroColuna4 = this.carregaFiltros();
    filtroColuna4.Sequencia = 9;
    filtroColuna4.Marca = new Array<PadraoComboFiltro>();
    filtroColuna4.Marca.push(marca);
    filtroColuna4.Onda = new Array<PadraoComboFiltro>();
    filtroColuna4.Onda.push(this.ondaColuna4);
    filtroColuna4.ParamDenominators = 1;// this.filtroService.ModelTarget.IdItem;
    this.dashBoardTwoService.carregarGraficoComparativoMarcas(filtroColuna4)
      .subscribe((coluna4: GraficoColunaModel) => {
        this.grafico2ColunaModel4 = coluna4

        this.ativaGraficoEvolutivoMarcas = true;

        // if (this.filtroService.ModelTarget.IdItem == 1) {
        // this.ativaGraficoEvolutivoMarcas = true;
        // }

        // if (this.filtroService.ModelTarget.IdItem == 2) {
        // this.ativaGraficoEvolutivoMarcasDenominators = true;
        // }

      }, (error) => console.error(error),
        () => { })
  }

  onchangeMarca(item: PadraoComboFiltro, nCol: number) {

    var filtroColuna1 = this.carregaFiltros();
    filtroColuna1.Marca = new Array<PadraoComboFiltro>();
    filtroColuna1.Marca.push(item);
    filtroColuna1.Sequencia = nCol;
    filtroColuna1.ParamDenominators = 1;// this.filtroService.ModelTarget.IdItem;
    this.dashBoardTwoService.carregarGraficoComparativoMarcas(filtroColuna1)
      .subscribe((coluna: GraficoColunaModel) => {

        switch (nCol) {
          case 1:
            this.graficoColunaModel1 = coluna
            this.marcaColuna1 = item;
            break;
          case 2:
            this.graficoColunaModel2 = coluna
            this.marcaColuna2 = item;
            break;
          case 3:
            this.graficoColunaModel3 = coluna
            this.marcaColuna3 = item;
            break;
          case 4:
            this.graficoColunaModel4 = coluna
            this.marcaColuna4 = item;
            break;
          case 5:
            this.graficoColunaModel5 = coluna
            this.marcaColuna5 = item;
            break;
        }

      }, (error) => console.error(error),
        () => { })


  }

  onchangeOnda(item: PadraoComboFiltro, nCol: number) {

    var filtroColuna1 = this.carregaFiltros();
    filtroColuna1.Marca = new Array<PadraoComboFiltro>();
    filtroColuna1.Marca.push(this.marcaEvolutivo);
    filtroColuna1.Onda = new Array<PadraoComboFiltro>();
    filtroColuna1.Onda.push(item);
    filtroColuna1.Sequencia = nCol;
    filtroColuna1.ParamDenominators = 1;// this.filtroService.ModelTarget.IdItem;
    this.dashBoardTwoService.carregarGraficoComparativoMarcas(filtroColuna1)
      .subscribe((coluna: GraficoColunaModel) => {


        switch (nCol) {
          case 1:
            this.grafico2ColunaModel1 = coluna
            this.ondaColuna1 = item;
            break;
          case 2:
            this.grafico2ColunaModel2 = coluna
            this.ondaColuna2 = item;
            break;
          case 3:
            this.grafico2ColunaModel3 = coluna
            this.ondaColuna3 = item;
            break;
          case 4:
            this.grafico2ColunaModel4 = coluna
            this.ondaColuna4 = item;
            break;
            // case 5:
            //   this.grafico2ColunaModel5 = coluna
            break;
        }

      }, (error) => console.error(error),
        () => { })


  }

  onchangeMarcaEvolutivo(item: PadraoComboFiltro) {
    this.carregarGraficoEvolutivoMarcasFirstLoad(item);
  }

  carregarGraficoComparativoMarcas() {
    this.dashBoardTwoService.carregarGraficoComparativoMarcas(null)
      .subscribe((response: GraficoColunaModel) => {
        this.graficoColunaModel1 = response
      }, (error) => console.error(error),
        () => { })
  }

  ajustaScrollFinal() {
    if (document.getElementById('aqui_id_div_grafico') != null && !document.getElementById('aqui_id_div_grafico').scrollLeft) {
      document.getElementById('aqui_id_div_grafico').scrollLeft = 9999;
    }

  }

  async DownloadExcelTabela() {
    await this.downloadArquivoService.DownloadExcelTabelaDinamicaExcelLojas()
      .subscribe((result) => {
        saveAs(
          result,
          "Nome_Expotação " + this.downloadArquivoService.getData() + ".xlsx"
        );
      }, (error) => console.error(error));
  }

  // exportToPptComparativoExperiencia() {
  //   this.conversorPowerpointService.captureScreenPPTAlternative(
  //     'comparativo-experiencia',
  //     "NPS comparativo por Experiencia",
  //   );
  // }

  exportToPptComparativoMarcas() {

    var nome = this.translate.instant('dashboard-two-grafico1-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'comparativo-marcas',
      nome,
    );
  }

  exportToPptEvolutivoMarcas() {
    var nome = this.translate.instant('dashboard-two-grafico2-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'evolutivo-marcas',
      nome
    );
  }



  async exportToExcelGrafico1() {

    var filtroPadrao = this.carregaFiltrosExcel();
    var titulo = this.translate.instant('dashboard-two-grafico1-titulo');
    filtroPadrao.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardTwoComparativoMarcas(filtroPadrao)
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
    var filtroPadrao = this.carregaFiltrosExcel();
    var titulo = this.translate.instant('dashboard-two-grafico2-titulo');
    filtroPadrao.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardTwoEvolutivoMarcas(filtroPadrao)
      .subscribe((result) => {

        saveAs(
          result,
          titulo + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => console.error(error));
  }




  async exportToExcelGrafico1Denominators() {

    var filtroPadrao = this.carregaFiltrosExcel();
    var titulo = this.translate.instant('dashboard-two-grafico1-titulo');
    filtroPadrao.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardTwoComparativoMarcasDenominators(filtroPadrao)
      .subscribe((result) => {

        saveAs(
          result,
          titulo + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => console.error(error));
  }

  async exportToExcelGrafico2Denominators() {
    var filtroPadrao = this.carregaFiltrosExcel();
    var titulo = this.translate.instant('dashboard-two-grafico2-titulo');
    filtroPadrao.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardTwoEvolutivoMarcasDenominators(filtroPadrao)
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
