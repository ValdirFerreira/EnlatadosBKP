import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import * as Highcharts from 'highcharts';
import { FiltroPadrao, FiltroPadraoExcel } from 'src/app/models/Filtros/FiltroPadrao';
import { saveAs } from "file-saver";
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';
import { MenuService } from 'src/app/services/menu.service';
import { DownloadArquivoService } from 'src/app/services/download-arquivo.service';
import { EventEmitterService } from 'src/app/services/event-emitter.service';
import { TranslateService } from '@ngx-translate/core';
import { ConversorPowerpointService } from 'src/app/services/conversor-powerpoint.service';

import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';
import { AuthService } from 'src/app/services/auth.service';
import { Session } from '../home/guards/session';
import { DashBoardThreeService } from 'src/app/services/dashboard-three.service';
import { GraficoFunil } from 'src/app/models/grafico-funil/GraficoFunil';
import { LogService } from 'src/app/services/log.service';



@Component({
  selector: 'app-dashboard-three',
  templateUrl: './dashboard-three.component.html',
  styleUrls: ['./dashboard-three.component.scss']
})
export class DashboardThreeComponent implements OnInit {

  graficoFunil1 = new GraficoFunil();
  graficoFunil2 = new GraficoFunil();
  graficoFunil3 = new GraficoFunil();
  graficoFunil4 = new GraficoFunil();
  graficoFunil5 = new GraficoFunil();
  graficoFunil6 = new GraficoFunil();
  graficoFunil7 = new GraficoFunil();
  graficoFunil8 = new GraficoFunil();


  grafico2Funil1 = new GraficoFunil();
  grafico2Funil2 = new GraficoFunil();
  grafico2Funil3 = new GraficoFunil();
  grafico2Funil4 = new GraficoFunil();

  // Filtros de Marca para utilização da geração do Excel Grafico Comparativo Marcas 
  marcaColuna1: PadraoComboFiltro;
  marcaColuna2: PadraoComboFiltro;
  marcaColuna3: PadraoComboFiltro;
  marcaColuna4: PadraoComboFiltro;
  marcaColuna5: PadraoComboFiltro;
  marcaColuna6: PadraoComboFiltro;
  marcaColuna7: PadraoComboFiltro;
  marcaColuna8: PadraoComboFiltro;
  // Filtros de Marca para utilização da geração do Excel Grafico Comparativo Marcas 

  // Filtros de onde para utilização da geração do Excel Grafico Marcas Evolutivo
  ondaColuna1: PadraoComboFiltro;
  ondaColuna2: PadraoComboFiltro;
  ondaColuna3: PadraoComboFiltro;
  ondaColuna4: PadraoComboFiltro;
  marcaEvolutivo: PadraoComboFiltro;
  // Filtros de onde para utilização da geração do Excel Grafico Marcas Evolutivo

  constructor(public router: Router,
    public menuService: MenuService,
    public filtroService: FiltroGlobalService, private downloadArquivoService: DownloadArquivoService,
    private translate: TranslateService, private conversorPowerpointService: ConversorPowerpointService,
    private authService: AuthService,
    private session: Session,
    private dashBoardThreeService: DashBoardThreeService,
    private logService: LogService,
  ) { }

  ativaGraficoComparativoMarcas: boolean = false;
  ativaGraficoEvolutivoMarcas: boolean = false;

  ngOnInit(): void {

    window.scroll({
      top: 0,
      left: 0,
      behavior: 'smooth'
    });

    this.menuService.nomePage = this.translate.instant('navbar.dashboard-three');

    EventEmitterService.get("emit-dashboard-three").subscribe((x) => {

      this.carregarGraficos();

      this.logService.GravaLogRota(this.router.url).subscribe(
      );

    })

    this.menuService.activeMenu = 1;
    this.menuService.menuSelecao = "1"

  }


  carregarGraficos() {

    this.filtroService.FiltroMarcas(this.filtroService.ModelRegiao)
      .subscribe((response: Array<PadraoComboFiltro>) => {
        this.filtroService.listaMarcas = response;

        this.ativaGraficoComparativoMarcas = false;
        this.ativaGraficoEvolutivoMarcas = false;

        this.marcaColuna1= null;
        this.marcaColuna2= null;
        this.marcaColuna3= null;
        this.marcaColuna4= null;
        this.marcaColuna5= null;
        this.marcaColuna6= null;
        this.marcaColuna7= null;
        this.marcaColuna8= null;

        this.marcaEvolutivo = null;


        this.ondaColuna1 = null;
        this.ondaColuna2 = null;
        this.ondaColuna3 = null;
        this.ondaColuna4 = null;




        this.carregaFiltroMarcas();
        this.carregaFiltroOndas();
        this.carregarGraficoComparativoMarcasFirstLoad();
        this.carregarGraficoEvolutivoMarcasFirstLoad(this.marcaEvolutivo);

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
    filtros.Onda = new Array<PadraoComboFiltro>();
    filtros.Onda.push(this.filtroService.ModelOnda);
    filtros.CodUser = this.session.getCodUserSession();
    filtros.CodIdioma = 1;// this.authService.idDefaultLangUser;
    filtros.ParamDenominators = this.filtroService.ModelDenominators.IdItem;
    return filtros;
  }


  public carregaFiltrosExcel() {

    var filtros = new FiltroPadraoExcel();
    var list =[];
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

    filtros.Marca6.push(this.marcaColuna6);
    filtros.Marca7.push(this.marcaColuna7);
    filtros.Marca8.push(this.marcaColuna8);

    filtros.CodUser = this.session.getCodUserSession();
    filtros.CodIdioma = 1;// this.authService.idDefaultLangUser;
    filtros.ParamDenominators = this.filtroService.ModelDenominators.IdItem;
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

    if (!this.marcaColuna6)
      this.marcaColuna6 = this.filtroService.listaMarcas[5]

    if (!this.marcaColuna7)
      this.marcaColuna7 = this.filtroService.listaMarcas[6]

    if (!this.marcaColuna8)
      this.marcaColuna8 = this.filtroService.listaMarcas[7]


    // utilizada no segundo grafico de funil
    if (!this.marcaEvolutivo)
      this.marcaEvolutivo = this.filtroService.listaMarcas[0]

  }

  public carregaFiltroOndas() {
    if (!this.ondaColuna1) {
      this.ondaColuna1 = this.filtroService.listaOnda[3];
    }

    if (!this.ondaColuna2) {
      this.ondaColuna2 = this.filtroService.listaOnda[2];
    }

    if (!this.ondaColuna3) {
      this.ondaColuna3 =  this.filtroService.listaOnda[1];
    }

    if (!this.ondaColuna4) {
      this.ondaColuna4 =  this.filtroService.listaOnda[0];
    }

  }

  carregarGraficoComparativoMarcasFirstLoad() {



    debugger

    var filtroColuna1 = this.carregaFiltros();
    filtroColuna1.Marca = new Array<PadraoComboFiltro>();
    filtroColuna1.Marca.push(this.marcaColuna1);
    filtroColuna1.Sequencia = 1;
    filtroColuna1.ParamDenominators = this.filtroService.ModelDenominators.IdItem;
    this.dashBoardThreeService.carregarGraficoComparativoExperiencia(filtroColuna1)
      .subscribe((coluna1: GraficoFunil) => {
        debugger
        this.graficoFunil1 = coluna1

      }, (error) => console.error(error),
        () => { })

    var filtroColuna2 = this.carregaFiltros();
    filtroColuna2.Sequencia = 2;
    filtroColuna2.Marca = new Array<PadraoComboFiltro>();
    filtroColuna2.Marca.push(this.marcaColuna2);
    filtroColuna2.ParamDenominators = this.filtroService.ModelDenominators.IdItem;
    this.dashBoardThreeService.carregarGraficoComparativoExperiencia(filtroColuna2)
      .subscribe((coluna2: GraficoFunil) => {
        this.graficoFunil2 = coluna2
      }, (error) => console.error(error),
        () => { })

    var filtroColuna3 = this.carregaFiltros();
    filtroColuna3.Sequencia = 3;
    filtroColuna3.Marca = new Array<PadraoComboFiltro>();
    filtroColuna3.Marca.push(this.marcaColuna3);
    filtroColuna3.ParamDenominators = this.filtroService.ModelDenominators.IdItem;
    this.dashBoardThreeService.carregarGraficoComparativoExperiencia(filtroColuna3)
      .subscribe((coluna3: GraficoFunil) => {
        this.graficoFunil3 = coluna3

      }, (error) => console.error(error),
        () => { })

    var filtroColuna4 = this.carregaFiltros();
    filtroColuna4.Sequencia = 4;
    filtroColuna4.Marca = new Array<PadraoComboFiltro>();
    filtroColuna4.Marca.push(this.marcaColuna4);
    filtroColuna4.ParamDenominators = this.filtroService.ModelDenominators.IdItem;
    this.dashBoardThreeService.carregarGraficoComparativoExperiencia(filtroColuna4)
      .subscribe((coluna4: GraficoFunil) => {
        this.graficoFunil4 = coluna4

      }, (error) => console.error(error),
        () => { })

    var filtroColuna5 = this.carregaFiltros();
    filtroColuna5.Sequencia = 5;
    filtroColuna5.Marca = new Array<PadraoComboFiltro>();
    filtroColuna5.Marca.push(this.marcaColuna5);
    filtroColuna5.ParamDenominators = this.filtroService.ModelDenominators.IdItem;
    this.dashBoardThreeService.carregarGraficoComparativoExperiencia(filtroColuna5)
      .subscribe((coluna5: GraficoFunil) => {
        this.graficoFunil5 = coluna5

      }, (error) => console.error(error),
        () => { })

    var filtroColuna6 = this.carregaFiltros();
    filtroColuna6.Sequencia = 6;
    filtroColuna6.Marca = new Array<PadraoComboFiltro>();
    filtroColuna6.Marca.push(this.marcaColuna6);
    filtroColuna6.ParamDenominators = this.filtroService.ModelDenominators.IdItem;
    this.dashBoardThreeService.carregarGraficoComparativoExperiencia(filtroColuna6)
      .subscribe((coluna6: GraficoFunil) => {
        this.graficoFunil6 = coluna6

      }, (error) => console.error(error),
        () => { })

    var filtroColuna7 = this.carregaFiltros();
    filtroColuna7.Sequencia = 7;
    filtroColuna7.Marca = new Array<PadraoComboFiltro>();
    filtroColuna7.Marca.push(this.marcaColuna7);
    filtroColuna7.ParamDenominators = this.filtroService.ModelDenominators.IdItem;
    this.dashBoardThreeService.carregarGraficoComparativoExperiencia(filtroColuna7)
      .subscribe((coluna7: GraficoFunil) => {
        this.graficoFunil7 = coluna7

      }, (error) => console.error(error),
        () => { })

    var filtroColuna8 = this.carregaFiltros();
    filtroColuna8.Sequencia = 8;
    filtroColuna8.Marca = new Array<PadraoComboFiltro>();
    filtroColuna8.Marca.push(this.marcaColuna8);
    filtroColuna8.ParamDenominators = this.filtroService.ModelDenominators.IdItem;
    this.dashBoardThreeService.carregarGraficoComparativoExperiencia(filtroColuna8)
      .subscribe((coluna8: GraficoFunil) => {
        this.graficoFunil8 = coluna8
        this.ativaGraficoComparativoMarcas = true;
      }, (error) => console.error(error),
        () => { })
  }

  carregarGraficoEvolutivoMarcasFirstLoad(item: PadraoComboFiltro = null) {



    var filtroColuna1 = this.carregaFiltros();
    filtroColuna1.Sequencia = 6;
    filtroColuna1.Marca = new Array<PadraoComboFiltro>();
    filtroColuna1.Marca.push(this.marcaEvolutivo);
    filtroColuna1.Onda = new Array<PadraoComboFiltro>();
    filtroColuna1.Onda.push(this.ondaColuna1);
    filtroColuna1.ParamDenominators = this.filtroService.ModelDenominators.IdItem;
    this.dashBoardThreeService.carregarGraficoComparativoExperiencia(filtroColuna1)
      .subscribe((coluna1: GraficoFunil) => {
        this.grafico2Funil1 = coluna1

      }, (error) => console.error(error),
        () => { })

    var filtroColuna2 = this.carregaFiltros();
    filtroColuna2.Sequencia = 7;
    filtroColuna2.Marca = new Array<PadraoComboFiltro>();
    filtroColuna2.Marca.push(this.marcaEvolutivo);
    filtroColuna2.Onda = new Array<PadraoComboFiltro>();
    filtroColuna2.Onda.push(this.ondaColuna2);
    filtroColuna2.ParamDenominators = this.filtroService.ModelDenominators.IdItem;
    this.dashBoardThreeService.carregarGraficoComparativoExperiencia(filtroColuna2)
      .subscribe((coluna2: GraficoFunil) => {
        this.grafico2Funil2 = coluna2

      }, (error) => console.error(error),
        () => { })

    var filtroColuna3 = this.carregaFiltros();
    filtroColuna3.Sequencia = 8;
    filtroColuna3.Marca = new Array<PadraoComboFiltro>();
    filtroColuna3.Marca.push(this.marcaEvolutivo);
    filtroColuna3.Onda = new Array<PadraoComboFiltro>();
    filtroColuna3.Onda.push(this.ondaColuna3);
    filtroColuna3.ParamDenominators = this.filtroService.ModelDenominators.IdItem;
    this.dashBoardThreeService.carregarGraficoComparativoExperiencia(filtroColuna3)
      .subscribe((coluna3: GraficoFunil) => {
        this.grafico2Funil3 = coluna3

      }, (error) => console.error(error),
        () => { })

    var filtroColuna4 = this.carregaFiltros();
    filtroColuna4.Sequencia = 9;
    filtroColuna4.Marca = new Array<PadraoComboFiltro>();
    filtroColuna4.Marca.push(this.marcaEvolutivo);
    filtroColuna4.Onda = new Array<PadraoComboFiltro>();
    filtroColuna4.Onda.push(this.ondaColuna4);
    filtroColuna4.ParamDenominators = this.filtroService.ModelDenominators.IdItem;
    this.dashBoardThreeService.carregarGraficoComparativoExperiencia(filtroColuna4)
      .subscribe((coluna4: GraficoFunil) => {
        this.grafico2Funil4 = coluna4
        this.ativaGraficoEvolutivoMarcas = true;
      }, (error) => console.error(error),
        () => { })
  }

  onchangeMarca(item: PadraoComboFiltro, nCol: number) {

    var filtroColuna1 = this.carregaFiltros();
    filtroColuna1.Marca = new Array<PadraoComboFiltro>();
    filtroColuna1.Marca.push(item);
    filtroColuna1.Sequencia = nCol;
    filtroColuna1.ParamDenominators = this.filtroService.ModelDenominators.IdItem;
    this.dashBoardThreeService.carregarGraficoComparativoExperiencia(filtroColuna1)
      .subscribe((coluna: GraficoFunil) => {

        switch (nCol) {
          case 1:
            this.graficoFunil1 = coluna
            this.marcaColuna1 = item;
            break;
          case 2:
            this.graficoFunil2 = coluna
            this.marcaColuna2 = item;
            break;
          case 3:
            this.graficoFunil3 = coluna
            this.marcaColuna3 = item;
            break;
          case 4:
            this.graficoFunil4 = coluna
            this.marcaColuna4 = item;
            break;
          case 5:
            this.graficoFunil5 = coluna
            this.marcaColuna5 = item;
            break;
          case 6:
            this.graficoFunil6 = coluna
            this.marcaColuna6 = item;
            break;
          case 7:
            this.graficoFunil7 = coluna
            this.marcaColuna7 = item;
            break;
          case 8:
            this.graficoFunil8 = coluna
            this.marcaColuna8 = item;
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
    filtroColuna1.ParamDenominators = this.filtroService.ModelDenominators.IdItem;
    this.dashBoardThreeService.carregarGraficoComparativoExperiencia(filtroColuna1)
      .subscribe((coluna: GraficoFunil) => {

        switch (nCol) {
          case 1:
            this.grafico2Funil1 = coluna
            this.ondaColuna1 = item;
            break;
          case 2:
            this.grafico2Funil2 = coluna
            this.ondaColuna2 = item;
            break;
          case 3:
            this.grafico2Funil3 = coluna
            this.ondaColuna3 = item;
            break;
          case 4:
            this.grafico2Funil4 = coluna
            this.ondaColuna4 = item;
            break;

        }

      }, (error) => console.error(error),
        () => { })


  }

  onchangeMarcaEvolutivo(item: PadraoComboFiltro) {
    this.marcaEvolutivo = item;
    this.carregarGraficoEvolutivoMarcasFirstLoad(item);
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


  async exportToExcelGrafico1() {

    var filtroPadrao = this.carregaFiltrosExcel();
    var titulo = this.translate.instant('dashboard-three-grafico1-titulo');
    filtroPadrao.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardThreeComparativoMarcas(filtroPadrao)
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
      DownloadDashboardThreeEvolutivoMarcas(filtroPadrao)
      .subscribe((result) => {

        saveAs(
          result,
          titulo + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => console.error(error));
  }

  exportToPptComparativoExperiencia() {
    var nome = this.translate.instant('dashboard-three-grafico1-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'comparativo-experiencia',
      nome,
    );
  }

  exportToPptEvolutivoMarcas() {
    var nome = this.translate.instant('dashboard-three-grafico2-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'evolutivo-marcas',
      nome
    );
  }

}
