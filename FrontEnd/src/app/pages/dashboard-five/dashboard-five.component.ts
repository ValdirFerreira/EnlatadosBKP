import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import * as Highcharts from 'highcharts';
import { FiltroPadrao, FiltroPadraoExcel, FiltroPadraoExcelDuplo } from 'src/app/models/Filtros/FiltroPadrao';
import { saveAs } from "file-saver";
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';
import { MenuService } from 'src/app/services/menu.service';
import { DownloadArquivoService } from 'src/app/services/download-arquivo.service';
import { EventEmitterService } from 'src/app/services/event-emitter.service';
import { GraficoColunaModel } from 'src/app/models/grafico-coluna/grafico-coluna';
import { TranslateService } from '@ngx-translate/core';
import { ConversorPowerpointService } from 'src/app/services/conversor-powerpoint.service';
import { AuthService } from 'src/app/services/auth.service';
import { Session } from '../home/guards/session';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';
import { DashBoardFiveService } from 'src/app/services/dashboard-five.service';
import { GraficoImagemPosicionamento } from 'src/app/models/grafico-Imagem-posicionamento/GraficoImagemPosicionamento';
import { LogService } from 'src/app/services/log.service';




@Component({
  selector: 'app-dashboard-five',
  templateUrl: './dashboard-five.component.html',
  styleUrls: ['./dashboard-five.component.scss']
})
export class DashboardFiveComponent implements OnInit {


  graficoColunaModel1 = new Array<GraficoImagemPosicionamento>();
  graficoColunaModel2 = new Array<GraficoImagemPosicionamento>();
  graficoColunaModel3 = new Array<GraficoImagemPosicionamento>();
  graficoColunaModel4 = new Array<GraficoImagemPosicionamento>();
  graficoColunaModel5 = new Array<GraficoImagemPosicionamento>();

  grafico2ColunaModel1 = new Array<GraficoImagemPosicionamento>();
  grafico2ColunaModel2 = new Array<GraficoImagemPosicionamento>();
  grafico2ColunaModel3 = new Array<GraficoImagemPosicionamento>();
  grafico2ColunaModel4 = new Array<GraficoImagemPosicionamento>();
  grafico2ColunaModel5 = new Array<GraficoImagemPosicionamento>();

  grafico3ColunaModel1 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel2 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel1_3 = new Array<GraficoImagemPosicionamento>();

  grafico3ColunaModel3 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel4 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel3_3 = new Array<GraficoImagemPosicionamento>();

  grafico3ColunaModel5 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel6 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel5_3 = new Array<GraficoImagemPosicionamento>();

  grafico3ColunaModel7 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel8 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel7_3 = new Array<GraficoImagemPosicionamento>();

  grafico3ColunaModel9 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel10 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel9_3 = new Array<GraficoImagemPosicionamento>();

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


  // Filtros de Marca para utilização da geração do Excel Grafico Comparativo Marcas 
  ondaDuploColuna1: PadraoComboFiltro;
  ondaDuploColuna1_2: PadraoComboFiltro;
  ondaDuploColuna1_3: PadraoComboFiltro;

  ondaDuploColuna2: PadraoComboFiltro;
  ondaDuploColuna2_2: PadraoComboFiltro;
  ondaDuploColuna2_3: PadraoComboFiltro;

  ondaDuploColuna3: PadraoComboFiltro;
  ondaDuploColuna3_2: PadraoComboFiltro;
  ondaDuploColuna3_3: PadraoComboFiltro;

  ondaDuploColuna4: PadraoComboFiltro;
  ondaDuploColuna4_2: PadraoComboFiltro;
  ondaDuploColuna4_3: PadraoComboFiltro;

  ondaDuploColuna5: PadraoComboFiltro;
  ondaDuploColuna5_2: PadraoComboFiltro;
  ondaDuploColuna5_3: PadraoComboFiltro;

  // Filtros de Marca para utilização da geração do Excel Grafico Comparativo Marcas 
  marcaDuploColuna1: PadraoComboFiltro;
  marcaDuploColuna2: PadraoComboFiltro;
  marcaDuploColuna3: PadraoComboFiltro;
  marcaDuploColuna4: PadraoComboFiltro;
  marcaDuploColuna5: PadraoComboFiltro;
  // Filtros de Marca para utilização da geração do Excel Grafico Comparativo Marcas 

  // Filtros de Marca para utilização da geração do Excel Grafico Comparativo Marcas 


  ativaGraficoComparativoMarcas: boolean = false;
  ativaGraficoEvolutivoMarcas: boolean = false;
  ativaGraficoComparativoMarcasDuplo: boolean = false;

  paginaAtiva: boolean = true;

  ajusteTelaFiltro: number = 1;

  constructor(public router: Router,
    public menuService: MenuService,
    public filtroService: FiltroGlobalService, private downloadArquivoService: DownloadArquivoService, private dashBoardFiveService: DashBoardFiveService,
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

    this.menuService.nomePage = this.translate.instant('navbar.dashboard-five');
    this.menuService.activeMenu = 5;
    this.menuService.menuSelecao = "5"

    EventEmitterService.get("emit-dashboard-five").subscribe((x) => {

      this.paginaAtiva = false;

      this.ativaGraficoComparativoMarcas = false;

      this.carregarGraficos();

      this.logService.GravaLogRota(this.router.url).subscribe(
      );

    })

  }






  carregarGraficos() {


    // if (this.filtroService.listaTarget)
    //   this.filtroService.ModelTarget = this.filtroService.listaTarget[0];
    // else {
    //   var item = new PadraoComboFiltro();
    //   item.IdItem = 1;
    //   this.filtroService.ModelTarget = item;
    // }


    this.filtroService.FiltroMarcas(this.filtroService.ModelRegiao)
      .subscribe((response: Array<PadraoComboFiltro>) => {
        this.filtroService.listaMarcas = response;



        this.marcaColuna1 = null;
        this.marcaColuna2 = null;
        this.marcaColuna3 = null;
        this.marcaColuna4 = null;
        this.marcaColuna5 = null;

        this.marcaDuploColuna1 = null;
        this.marcaDuploColuna2 = null;
        this.marcaDuploColuna3 = null;
        this.marcaDuploColuna4 = null;
        this.marcaDuploColuna5 = null;

        this.ondaColuna1 = null;
        this.ondaColuna2 = null;
        this.ondaColuna3 = null;
        this.ondaColuna4 = null;

        this.ondaDuploColuna1 = null;
        this.ondaDuploColuna1_2 = null;
        this.ondaDuploColuna1_3 = null;

        this.ondaDuploColuna2 = null;
        this.ondaDuploColuna2_2 = null;
        this.ondaDuploColuna2_3 = null;

        this.ondaDuploColuna3 = null;
        this.ondaDuploColuna3_2 = null;
        this.ondaDuploColuna3_3 = null;

        this.ondaDuploColuna4 = null;
        this.ondaDuploColuna4_2 = null;
        this.ondaDuploColuna4_3 = null;

        this.ondaDuploColuna5 = null;
        this.ondaDuploColuna5_2 = null;
        this.ondaDuploColuna5_3 = null;

        this.marcaColuna1 = this.filtroService.listaMarcas[0]
        this.marcaColuna2 = this.filtroService.listaMarcas[1]
        this.marcaColuna3 = this.filtroService.listaMarcas[2]
        this.marcaColuna4 = this.filtroService.listaMarcas[3]
        this.marcaColuna5 = this.filtroService.listaMarcas[4]

        // Marcas grafico duplo
        this.marcaDuploColuna1 = this.filtroService.listaMarcas[0]
        this.marcaDuploColuna2 = this.filtroService.listaMarcas[1]
        this.marcaDuploColuna3 = this.filtroService.listaMarcas[2]
        this.marcaDuploColuna4 = this.filtroService.listaMarcas[3]
        this.marcaDuploColuna5 = this.filtroService.listaMarcas[4]



        // this.ajusteTelaFiltro = this.filtroService.ModelTarget.IdItem;
        this.paginaAtiva = true;

        this.carregaFiltroMarcas();
        this.carregaFiltroOndas();
        this.carregarGraficoComparativoMarcasFirstLoad();
        // this.carregarGraficoEvolutivoMarcasFirstLoad(this.marcaEvolutivo);
        this.carregarGraficoEvolutivoMarcasFirstLoad(null);
        this.carregarGraficoEvolutivoMarcasDuploFirstLoad();

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
    return filtros;
  }


  carregarGraficoComparativoMarcasFirstLoad() {


   
    var filtroColuna1 = this.carregaFiltros();
    filtroColuna1.Marca = new Array<PadraoComboFiltro>();
    filtroColuna1.Marca.push(this.marcaColuna1);
    filtroColuna1.Sequencia = 1;

    this.dashBoardFiveService.carregarGraficoComparativoMarcas(filtroColuna1)
      .subscribe((coluna1: Array<GraficoImagemPosicionamento>) => {
        this.graficoColunaModel1 = coluna1
      }, (error) => { console.error(error) },
        () => { })

    var filtroColuna2 = this.carregaFiltros();
    filtroColuna2.Sequencia = 2;
    filtroColuna2.Marca = new Array<PadraoComboFiltro>();
    filtroColuna2.Marca.push(this.marcaColuna2);
    this.dashBoardFiveService.carregarGraficoComparativoMarcas(filtroColuna2)
      .subscribe((coluna2: Array<GraficoImagemPosicionamento>) => {
        this.graficoColunaModel2 = coluna2

      }, (error) => { console.error(error) },
        () => { })

    var filtroColuna3 = this.carregaFiltros();
    filtroColuna3.Sequencia = 3;
    filtroColuna3.Marca = new Array<PadraoComboFiltro>();
    filtroColuna3.Marca.push(this.marcaColuna3);
    this.dashBoardFiveService.carregarGraficoComparativoMarcas(filtroColuna3)
      .subscribe((coluna3: Array<GraficoImagemPosicionamento>) => {

        this.graficoColunaModel3 = coluna3

      }, (error) => { console.error(error) },
        () => { })

    var filtroColuna4 = this.carregaFiltros();
    filtroColuna4.Sequencia = 4;
    filtroColuna4.Marca = new Array<PadraoComboFiltro>();
    filtroColuna4.Marca.push(this.marcaColuna4);

    this.dashBoardFiveService.carregarGraficoComparativoMarcas(filtroColuna4)
      .subscribe((coluna4: Array<GraficoImagemPosicionamento>) => {
        this.graficoColunaModel4 = coluna4

      }, (error) => { console.error(error) },
        () => { })

    var filtroColuna5 = this.carregaFiltros();
    filtroColuna5.Sequencia = 5;
    filtroColuna5.Marca = new Array<PadraoComboFiltro>();
    filtroColuna5.Marca.push(this.marcaColuna5);
    this.dashBoardFiveService.carregarGraficoComparativoMarcas(filtroColuna5)
      .subscribe((coluna5: Array<GraficoImagemPosicionamento>) => {
        this.graficoColunaModel5 = coluna5
        this.ativaGraficoComparativoMarcas = true;
      }, (error) => { console.error(error) },
        () => { })

  }

  carregarGraficoEvolutivoMarcasFirstLoad(item: PadraoComboFiltro = null) {


    var marca = item == null ? this.filtroService.listaMarcas[0] : item;


    var itemDefault = new PadraoComboFiltro();
    itemDefault.DescItem = "";
    itemDefault.IdItem = 0;

    if (!this.ondaColuna1) {
      this.ondaColuna1 = itemDefault;
    }

    if (!this.ondaColuna2) {
      this.ondaColuna2 = itemDefault;
    }

    if (!this.ondaColuna3) {
      this.ondaColuna3 = itemDefault;
    }

    if (!this.ondaColuna3) {
      this.ondaColuna3 = itemDefault;
    }

    if (!this.ondaColuna4) {
      this.ondaColuna4 = itemDefault;
    }

    var filtroColuna1 = this.carregaFiltros();
    filtroColuna1.Sequencia = 6;
    filtroColuna1.Marca = new Array<PadraoComboFiltro>();
    filtroColuna1.Marca.push(marca);
    filtroColuna1.Onda = new Array<PadraoComboFiltro>();
    filtroColuna1.Onda.push(this.ondaColuna1);
    this.marcaEvolutivo = marca;


    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna1)
      .subscribe((coluna1: Array<GraficoImagemPosicionamento>) => {
        this.grafico2ColunaModel1 = coluna1

      }, (error) => console.error(error),
        () => { })

    var filtroColuna2 = this.carregaFiltros();
    filtroColuna2.Sequencia = 7;
    filtroColuna2.Marca = new Array<PadraoComboFiltro>();
    filtroColuna2.Marca.push(marca);
    filtroColuna2.Onda = new Array<PadraoComboFiltro>();
    filtroColuna2.Onda.push(this.ondaColuna2);

    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna2)
      .subscribe((coluna2: Array<GraficoImagemPosicionamento>) => {
        this.grafico2ColunaModel2 = coluna2

      }, (error) => console.error(error),
        () => { })

    var filtroColuna3 = this.carregaFiltros();
    filtroColuna3.Sequencia = 8;
    filtroColuna3.Marca = new Array<PadraoComboFiltro>();
    filtroColuna3.Marca.push(marca);
    filtroColuna3.Onda = new Array<PadraoComboFiltro>();
    filtroColuna3.Onda.push(this.ondaColuna3);

    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna3)
      .subscribe((coluna3: Array<GraficoImagemPosicionamento>) => {
        this.grafico2ColunaModel3 = coluna3

      }, (error) => console.error(error),
        () => { })

    var filtroColuna4 = this.carregaFiltros();
    filtroColuna4.Sequencia = 9;
    filtroColuna4.Marca = new Array<PadraoComboFiltro>();
    filtroColuna4.Marca.push(marca);
    filtroColuna4.Onda = new Array<PadraoComboFiltro>();
    filtroColuna4.Onda.push(this.ondaColuna4);

    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna4)
      .subscribe((coluna4: Array<GraficoImagemPosicionamento>) => {
        this.grafico2ColunaModel4 = coluna4
        this.ativaGraficoEvolutivoMarcas = true;
      }, (error) => console.error(error),
        () => { })
  }

  carregarGraficoEvolutivoMarcasDuploFirstLoad() {

    var itemDefault = new PadraoComboFiltro();
    itemDefault.DescItem = "";
    itemDefault.IdItem = 0;

    if (!this.ondaDuploColuna1) {
      this.ondaDuploColuna1 = itemDefault;
    }

    if (!this.ondaDuploColuna1_2) {
      this.ondaDuploColuna1_2 = itemDefault;
    }

    if (!this.ondaDuploColuna1_3) {
      this.ondaDuploColuna1_3 = itemDefault;
    }
    ///////////////////////


    if (!this.ondaDuploColuna2) {
      this.ondaDuploColuna2 = itemDefault;
    }

    if (!this.ondaDuploColuna2_2) {
      this.ondaDuploColuna2_2 = itemDefault;
    }

    if (!this.ondaDuploColuna2_3) {
      this.ondaDuploColuna2_3 = itemDefault;
    }
    ///////////////////////


    if (!this.ondaDuploColuna3) {
      this.ondaDuploColuna3 = itemDefault;
    }

    if (!this.ondaDuploColuna3_2) {
      this.ondaDuploColuna3_2 = itemDefault;
    }

    if (!this.ondaDuploColuna3_3) {
      this.ondaDuploColuna3_3 = itemDefault;
    }
    ///////////////////////

    if (!this.ondaDuploColuna4) {
      this.ondaDuploColuna4 = itemDefault;
    }

    if (!this.ondaDuploColuna4_2) {
      this.ondaDuploColuna4_2 = itemDefault;
    }

    if (!this.ondaDuploColuna4_3) {
      this.ondaDuploColuna4_3 = itemDefault;
    }

    ///////////////////////

    if (!this.ondaDuploColuna5) {
      this.ondaDuploColuna5 = itemDefault;
    }

    if (!this.ondaDuploColuna5_2) {
      this.ondaDuploColuna5_2 = itemDefault;
    }


    if (!this.ondaDuploColuna5_3) {
      this.ondaDuploColuna5_3 = itemDefault;
    }


    var filtroColuna1 = this.carregaFiltros();
    filtroColuna1.Sequencia = 7;
    filtroColuna1.Marca = new Array<PadraoComboFiltro>();
    filtroColuna1.Marca.push(this.marcaDuploColuna1);
    filtroColuna1.Onda = new Array<PadraoComboFiltro>();
    filtroColuna1.Onda.push(this.ondaDuploColuna1);
    debugger
    // Grafico 1
    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna1)
      .subscribe((coluna1: Array<GraficoImagemPosicionamento>) => {
        
        this.grafico3ColunaModel1 = coluna1

        debugger

      }, (error) => console.error(error),
        () => { })

    // Grafico 1
    var filtroColuna2 = this.carregaFiltros();
    filtroColuna2.Sequencia = 8;
    filtroColuna2.Marca = new Array<PadraoComboFiltro>();
    filtroColuna2.Marca.push(this.marcaDuploColuna1);
    filtroColuna2.Onda = new Array<PadraoComboFiltro>();
    filtroColuna2.Onda.push(this.ondaDuploColuna1_2);

    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna2)
      .subscribe((coluna2: Array<GraficoImagemPosicionamento>) => {
        this.grafico3ColunaModel2 = coluna2

      }, (error) => console.error(error),
        () => { })

    var filtroColuna1_3 = this.carregaFiltros();
    filtroColuna1_3.Sequencia = 73;
    filtroColuna1_3.Marca = new Array<PadraoComboFiltro>();
    filtroColuna1_3.Marca.push(this.marcaDuploColuna1);
    filtroColuna1_3.Onda = new Array<PadraoComboFiltro>();
    filtroColuna1_3.Onda.push(this.ondaDuploColuna1_3);

    // Grafico 1
    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna1_3)
      .subscribe((coluna1_3: Array<GraficoImagemPosicionamento>) => {
        this.grafico3ColunaModel1_3 = coluna1_3

      }, (error) => console.error(error),
        () => { })

    //////////////////////////////////////////////////////////////////



    // Grafico 2
    var filtroColuna3 = this.carregaFiltros();
    filtroColuna3.Sequencia = 9;
    filtroColuna3.Marca = new Array<PadraoComboFiltro>();
    filtroColuna3.Marca.push(this.marcaDuploColuna2);
    filtroColuna3.Onda = new Array<PadraoComboFiltro>();
    filtroColuna3.Onda.push(this.ondaDuploColuna2);

    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna3)
      .subscribe((coluna3: Array<GraficoImagemPosicionamento>) => {
        this.grafico3ColunaModel3 = coluna3

      }, (error) => console.error(error),
        () => { })

    // Grafico 2
    var filtroColuna4 = this.carregaFiltros();
    filtroColuna4.Sequencia = 10;
    filtroColuna4.Marca = new Array<PadraoComboFiltro>();
    filtroColuna4.Marca.push(this.marcaDuploColuna2);
    filtroColuna4.Onda = new Array<PadraoComboFiltro>();
    filtroColuna4.Onda.push(this.ondaDuploColuna2_2);

    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna4)
      .subscribe((coluna4: Array<GraficoImagemPosicionamento>) => {
        this.grafico3ColunaModel4 = coluna4

      }, (error) => console.error(error),
        () => { })

        var filtroColuna3_3 = this.carregaFiltros();
        filtroColuna3_3.Sequencia = 93;
        filtroColuna3_3.Marca = new Array<PadraoComboFiltro>();
        filtroColuna3_3.Marca.push(this.marcaDuploColuna2);
        filtroColuna3_3.Onda = new Array<PadraoComboFiltro>();
        filtroColuna3_3.Onda.push(this.ondaDuploColuna3_3);

    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna3_3)
      .subscribe((coluna3_3: Array<GraficoImagemPosicionamento>) => {
        this.grafico3ColunaModel3_3 = coluna3_3

      }, (error) => console.error(error),
        () => { })


        ///////////////////////////////////////////////////////

    // Grafico 3
    var filtroColuna5 = this.carregaFiltros();
    filtroColuna5.Sequencia = 11;
    filtroColuna5.Marca = new Array<PadraoComboFiltro>();
    filtroColuna5.Marca.push(this.marcaDuploColuna3);
    filtroColuna5.Onda = new Array<PadraoComboFiltro>();
    filtroColuna5.Onda.push(this.ondaDuploColuna3);

    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna5)
      .subscribe((coluna5: Array<GraficoImagemPosicionamento>) => {
        this.grafico3ColunaModel5 = coluna5

      }, (error) => console.error(error),
        () => { })

    // Grafico 3
    var filtroColuna6 = this.carregaFiltros();
    filtroColuna6.Sequencia = 12;
    filtroColuna6.Marca = new Array<PadraoComboFiltro>();
    filtroColuna6.Marca.push(this.marcaDuploColuna3);
    filtroColuna6.Onda = new Array<PadraoComboFiltro>();
    filtroColuna6.Onda.push(this.ondaDuploColuna3_2);

    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna6)
      .subscribe((coluna6: Array<GraficoImagemPosicionamento>) => {
        this.grafico3ColunaModel6 = coluna6

      }, (error) => console.error(error),
        () => { })

        var filtroColuna5_3 = this.carregaFiltros();
        filtroColuna5_3.Sequencia = 113;
        filtroColuna5_3.Marca = new Array<PadraoComboFiltro>();
        filtroColuna5_3.Marca.push(this.marcaDuploColuna3);
        filtroColuna5_3.Onda = new Array<PadraoComboFiltro>();
        filtroColuna5_3.Onda.push(this.ondaDuploColuna3_3);
    
        this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna5_3)
          .subscribe((coluna5_3: Array<GraficoImagemPosicionamento>) => {
            this.grafico3ColunaModel5_3 = coluna5_3
    
          }, (error) => console.error(error),
            () => { })

        ////////////////////////////////////////////////////////////////

    // Grafico 4
    var filtroColuna7 = this.carregaFiltros();
    filtroColuna7.Sequencia = 13;
    filtroColuna7.Marca = new Array<PadraoComboFiltro>();
    filtroColuna7.Marca.push(this.marcaDuploColuna4);
    filtroColuna7.Onda = new Array<PadraoComboFiltro>();
    filtroColuna7.Onda.push(this.ondaDuploColuna4);

    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna7)
      .subscribe((coluna7: Array<GraficoImagemPosicionamento>) => {
        this.grafico3ColunaModel7 = coluna7

      }, (error) => console.error(error),
        () => { })


    // Grafico 4
    var filtroColuna8 = this.carregaFiltros();
    filtroColuna8.Sequencia = 14;
    filtroColuna8.Marca = new Array<PadraoComboFiltro>();
    filtroColuna8.Marca.push(this.marcaDuploColuna4);
    filtroColuna8.Onda = new Array<PadraoComboFiltro>();
    filtroColuna8.Onda.push(this.ondaDuploColuna4_2);

    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna8)
      .subscribe((coluna8: Array<GraficoImagemPosicionamento>) => {
        this.grafico3ColunaModel8 = coluna8

      }, (error) => console.error(error),
        () => { })

        var filtroColuna7_3 = this.carregaFiltros();
        filtroColuna7_3.Sequencia = 133;
        filtroColuna7_3.Marca = new Array<PadraoComboFiltro>();
        filtroColuna7_3.Marca.push(this.marcaDuploColuna4);
        filtroColuna7_3.Onda = new Array<PadraoComboFiltro>();
        filtroColuna7_3.Onda.push(this.ondaDuploColuna4_3);
    
        this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna7_3)
          .subscribe((coluna7_3: Array<GraficoImagemPosicionamento>) => {
            this.grafico3ColunaModel7_3 = coluna7_3
    
          }, (error) => console.error(error),
            () => { })

        ////////////////////////////////////////


    // Grafico 5
    var filtroColuna9 = this.carregaFiltros();
    filtroColuna9.Sequencia = 15;
    filtroColuna9.Marca = new Array<PadraoComboFiltro>();
    filtroColuna9.Marca.push(this.marcaDuploColuna5);
    filtroColuna9.Onda = new Array<PadraoComboFiltro>();
    filtroColuna9.Onda.push(this.ondaDuploColuna5);

    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna9)
      .subscribe((coluna9: Array<GraficoImagemPosicionamento>) => {

        this.grafico3ColunaModel9 = coluna9

      }, (error) => console.error(error),
        () => { })

    // Grafico 5
    var filtroColuna10 = this.carregaFiltros();
    filtroColuna10.Sequencia = 16;
    filtroColuna10.Marca = new Array<PadraoComboFiltro>();
    filtroColuna10.Marca.push(this.marcaDuploColuna5);
    filtroColuna10.Onda = new Array<PadraoComboFiltro>();
    filtroColuna10.Onda.push(this.ondaDuploColuna5_2);

    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna10)
      .subscribe((coluna10: Array<GraficoImagemPosicionamento>) => {
        this.grafico3ColunaModel10 = coluna10
        this.ativaGraficoComparativoMarcasDuplo = true;
      }, (error) => console.error(error),
        () => { })



        var filtroColuna9_3 = this.carregaFiltros();
        filtroColuna9_3.Sequencia = 153;
        filtroColuna9_3.Marca = new Array<PadraoComboFiltro>();
        filtroColuna9_3.Marca.push(this.marcaDuploColuna5);
        filtroColuna9_3.Onda = new Array<PadraoComboFiltro>();
        filtroColuna9_3.Onda.push(this.ondaDuploColuna5_3);
    
        this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna9_3)
          .subscribe((coluna9_3: Array<GraficoImagemPosicionamento>) => {
    
            this.grafico3ColunaModel9_3 = coluna9_3
    
          }, (error) => console.error(error),
            () => { })



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


    // Marcas grafico duplo
    if (!this.marcaDuploColuna1)
      this.marcaDuploColuna1 = this.filtroService.listaMarcas[0]

    if (!this.marcaDuploColuna2)
      this.marcaDuploColuna2 = this.filtroService.listaMarcas[1]

    if (!this.marcaDuploColuna3)
      this.marcaDuploColuna3 = this.filtroService.listaMarcas[2]

    if (!this.marcaDuploColuna4)
      this.marcaDuploColuna4 = this.filtroService.listaMarcas[3]

    if (!this.marcaDuploColuna5)
      this.marcaDuploColuna5 = this.filtroService.listaMarcas[4]
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

    // Utilizada no grafico Duplo
    if (!this.ondaDuploColuna1) {
      this.ondaDuploColuna1 = this.filtroService.listaOnda[2];
    }
    if (!this.ondaDuploColuna1_2) {
      this.ondaDuploColuna1_2 = this.filtroService.listaOnda[1];
    }
    if (!this.ondaDuploColuna1_3) {
      this.ondaDuploColuna1_3 = this.filtroService.listaOnda[0];
    }

    if (!this.ondaDuploColuna2) {
      this.ondaDuploColuna2 = this.filtroService.listaOnda[2];
    }
    if (!this.ondaDuploColuna2_2) {
      this.ondaDuploColuna2_2 = this.filtroService.listaOnda[1];
    }
    if (!this.ondaDuploColuna2_3) {
      this.ondaDuploColuna2_3 = this.filtroService.listaOnda[0];
    }


    if (!this.ondaDuploColuna3) {
      this.ondaDuploColuna3 = this.filtroService.listaOnda[2];
    }
    if (!this.ondaDuploColuna3_2) {
      this.ondaDuploColuna3_2 = this.filtroService.listaOnda[1];
    }
    if (!this.ondaDuploColuna3_3) {
      this.ondaDuploColuna3_3 = this.filtroService.listaOnda[0];
    }


    if (!this.ondaDuploColuna4) {
      this.ondaDuploColuna4 = this.filtroService.listaOnda[2];
    }
    if (!this.ondaDuploColuna4_2) {
      this.ondaDuploColuna4_2 = this.filtroService.listaOnda[1];
    }
    if (!this.ondaDuploColuna4_3) {
      this.ondaDuploColuna4_3 = this.filtroService.listaOnda[0];
    }


    if (!this.ondaDuploColuna5) {
      this.ondaDuploColuna5 = this.filtroService.listaOnda[2];
    }
    if (!this.ondaDuploColuna5_2) {
      this.ondaDuploColuna5_2 = this.filtroService.listaOnda[1];
    }
    if (!this.ondaDuploColuna5_3) {
      this.ondaDuploColuna5_3 = this.filtroService.listaOnda[0];
    }


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

    filtros.CodUser = this.session.getCodUserSession();
    filtros.CodIdioma = 1;// this.authService.idDefaultLangUser;
    return filtros;
  }

  public carregaFiltrosExcelGraficoDuplo() {

    var filtros = new FiltroPadraoExcelDuplo();
    var list = [];
    list.push(this.filtroService.ModelTarget);
    filtros.Target = list;
    filtros.Regiao = this.filtroService.ModelRegiao;
    filtros.Demografico = this.filtroService.ModelDemografico;
    filtros.Onda.push(this.filtroService.ModelOnda);

    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Marca.push(this.marcaEvolutivo);


    filtros.OndaDuploColuna1 = this.ondaDuploColuna1;
    filtros.OndaDuploColuna1_2 = this.ondaDuploColuna1_2;
    filtros.OndaDuploColuna1_3 = this.ondaDuploColuna1_3;

    filtros.OndaDuploColuna2 = this.ondaDuploColuna2;
    filtros.OndaDuploColuna2_2 = this.ondaDuploColuna2_2;
    filtros.OndaDuploColuna2_3 = this.ondaDuploColuna2_3;

    filtros.OndaDuploColuna3 = this.ondaDuploColuna3;
    filtros.OndaDuploColuna3_2 = this.ondaDuploColuna3_2;
    filtros.OndaDuploColuna3_3 = this.ondaDuploColuna3_3;

    filtros.OndaDuploColuna4 = this.ondaDuploColuna4;
    filtros.OndaDuploColuna4_2 = this.ondaDuploColuna4_2;
    filtros.OndaDuploColuna4_3 = this.ondaDuploColuna4_3;

    filtros.OndaDuploColuna5 = this.ondaDuploColuna5;
    filtros.OndaDuploColuna5_2 = this.ondaDuploColuna5_2;
    filtros.OndaDuploColuna5_3 = this.ondaDuploColuna5_3;

    filtros.MarcaDuploColuna1 = this.marcaDuploColuna1;
    filtros.MarcaDuploColuna2 = this.marcaDuploColuna2;
    filtros.MarcaDuploColuna3 = this.marcaDuploColuna3;
    filtros.MarcaDuploColuna4 = this.marcaDuploColuna4;
    filtros.MarcaDuploColuna5 = this.marcaDuploColuna5;

    filtros.CodUser = this.session.getCodUserSession();
    filtros.CodIdioma = 1;// this.authService.idDefaultLangUser;
    return filtros;
  }



  // grafico Comparativo de marcas
  onchangeMarca(item: PadraoComboFiltro, nCol: number) {

    var filtroColuna = this.carregaFiltros();
    filtroColuna.Marca = new Array<PadraoComboFiltro>();
    filtroColuna.Marca.push(item);
    filtroColuna.Sequencia = nCol;
    this.dashBoardFiveService.carregarGraficoComparativoMarcas(filtroColuna)
      .subscribe((coluna: Array<GraficoImagemPosicionamento>) => {

        switch (nCol) {
          case 1:
            this.graficoColunaModel1 = coluna;
            this.marcaColuna1 = item;
            break;
          case 2:
            this.graficoColunaModel2 = coluna;
            this.marcaColuna2 = item;
            break;
          case 3:
            this.graficoColunaModel3 = coluna;
            this.marcaColuna3 = item;
            break;
          case 4:
            this.graficoColunaModel4 = coluna;
            this.marcaColuna4 = item;
            break;
          case 5:
            this.graficoColunaModel5 = coluna;
            this.marcaColuna5 = item;
            break;
        }

      }, (error) => { console.error(error) },
        () => { })


  }


  //Método grafico evolutivo
  onchangeEvolutivoOnda(item: PadraoComboFiltro, nCol: number) {

    var filtroColuna1 = this.carregaFiltros();
    filtroColuna1.Marca = new Array<PadraoComboFiltro>();
    filtroColuna1.Marca.push(this.marcaEvolutivo);
    filtroColuna1.Onda = new Array<PadraoComboFiltro>();
    filtroColuna1.Onda.push(item);
    filtroColuna1.Sequencia = nCol;

    this.dashBoardFiveService.carregarGraficoComparativoMarcas(filtroColuna1)
      .subscribe((coluna: Array<GraficoImagemPosicionamento>) => {


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


  // Métodos grafico Duplo
  onchangeOndaDuplo(item: PadraoComboFiltro, nCol: number) {

    var filtroColuna = this.carregaFiltros();
    filtroColuna.Marca = new Array<PadraoComboFiltro>();

    filtroColuna.Onda = new Array<PadraoComboFiltro>();
    filtroColuna.Onda.push(item);
    filtroColuna.Sequencia = nCol;
 
    switch (nCol) {
      case 1:
      case 13:
      case 2:
        filtroColuna.Marca.push(this.marcaDuploColuna1);
        break;

      case 3:
      case 33:
      case 4:
        filtroColuna.Marca.push(this.marcaDuploColuna2);
        break;

      case 5:
      case 53:
      case 6:
        filtroColuna.Marca.push(this.marcaDuploColuna3);
        break;

      case 7:
      case 73:
      case 8:
        filtroColuna.Marca.push(this.marcaDuploColuna4);
        break;

      case 9:
      case 93:
      case 10:
        filtroColuna.Marca.push(this.marcaDuploColuna5);
        break;

    }


    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna)
      .subscribe((coluna: Array<GraficoImagemPosicionamento>) => {


        switch (nCol) {
          case 1:
            this.grafico3ColunaModel1 = coluna;
            this.ondaDuploColuna1 = item;
            break;
          case 2:
            this.grafico3ColunaModel2 = coluna;
            this.ondaDuploColuna1_2 = item;
            break;
          case 13:
            this.grafico3ColunaModel1_3 = coluna;
            this.ondaDuploColuna1_3 = item;
            break;


          case 3:
            this.grafico3ColunaModel3 = coluna;
            this.ondaDuploColuna2 = item;
            break;
          case 4:
            this.grafico3ColunaModel4 = coluna;
            this.ondaDuploColuna2_2 = item;
            break;
          case 33:
            this.grafico3ColunaModel3_3 = coluna;
            this.ondaDuploColuna2_3 = item;
            break;

          case 5:
            this.grafico3ColunaModel5 = coluna;
            this.ondaDuploColuna3 = item;
            break;
          case 6:
            this.grafico3ColunaModel6 = coluna;
            this.ondaDuploColuna3_2 = item;
            break;
          case 53:
            this.grafico3ColunaModel5_3 = coluna;
            this.ondaDuploColuna3_3 = item;
            break;

          case 7:
            this.grafico3ColunaModel7 = coluna;
            this.ondaDuploColuna4 = item;
            break;
          case 8:
            this.grafico3ColunaModel8 = coluna;
            this.ondaDuploColuna4_2 = item;
            break;
          case 73:
            this.grafico3ColunaModel7_3 = coluna;
            this.ondaDuploColuna4_3 = item;
            break;

          case 9:
            this.grafico3ColunaModel9 = coluna;
            this.ondaDuploColuna5 = item;
            break;
          case 10:
            this.grafico3ColunaModel10 = coluna;
            this.ondaDuploColuna5_2 = item;
            break;
          case 93:
            this.grafico3ColunaModel9_3 = null;// coluna;
            this.ondaDuploColuna5_3 = item;
            break;
            // case 5:
            //   this.grafico2ColunaModel5 = coluna
            break;
        }

      }, (error) => { console.error(error) },
        () => { })


  }

  onchangeDuploMarca(item: PadraoComboFiltro, nCol: number) {

    this.carregaFiltroOndas()
    this.carregaFiltroMarcas()

    var onda1 = new PadraoComboFiltro();
    var onda2 = new PadraoComboFiltro();
    var onda3 = new PadraoComboFiltro();

    switch (nCol) {
      case 1:

        onda1 = this.ondaDuploColuna1;
        onda2 = this.ondaDuploColuna1_2;
        onda3 = this.ondaDuploColuna1_3;
        break;

      case 2:
        onda1 = this.ondaDuploColuna2;
        onda2 = this.ondaDuploColuna2_2;
        onda3 = this.ondaDuploColuna2_3;
        break;

      case 3:
        onda1 = this.ondaDuploColuna3;
        onda2 = this.ondaDuploColuna3_2;
        onda3 = this.ondaDuploColuna3_3;
        break;

      case 4:
        onda1 = this.ondaDuploColuna4;
        onda2 = this.ondaDuploColuna4_2;
        onda3 = this.ondaDuploColuna4_3;
        break;

      case 5:
        onda1 = this.ondaDuploColuna5;
        onda2 = this.ondaDuploColuna5_2;
        onda3 = this.ondaDuploColuna5_3;
        break;
    }



    // var filtroColuna1 = this.carregaFiltros();
    // filtroColuna1.Sequencia = 7;
    // filtroColuna1.Marca = new Array<PadraoComboFiltro>();
    // filtroColuna1.Marca.push(this.marcaDuploColuna1);
    // filtroColuna1.Onda = new Array<PadraoComboFiltro>();
    // filtroColuna1.Onda.push(this.ondaDuploColuna1);



    // var filtroColuna1 = this.carregaFiltros();
    // filtroColuna1.Sequencia = 7;
    // filtroColuna1.Marca = new Array<PadraoComboFiltro>();
    // filtroColuna1.Marca.push(this.marcaDuploColuna1);
    // filtroColuna1.Onda = new Array<PadraoComboFiltro>();
    // filtroColuna1.Onda.push(this.ondaDuploColuna1);
    // debugger


    var filtroColuna = this.carregaFiltros();
    filtroColuna.Marca = new Array<PadraoComboFiltro>();
    filtroColuna.Marca.push(item);
    filtroColuna.Sequencia = 7;
    filtroColuna.Onda = new Array<PadraoComboFiltro>();
    filtroColuna.Onda.push(onda1);
    debugger
    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColuna)
      .subscribe((coluna: Array<GraficoImagemPosicionamento>) => {

        switch (nCol) {
          case 1:

          debugger
            this.marcaDuploColuna1 = item;
            this.grafico3ColunaModel1 = null
            this.grafico3ColunaModel1 = coluna;
            break;
          case 2:

            this.marcaDuploColuna2 = item;
            this.grafico3ColunaModel3 = null;
            this.grafico3ColunaModel3 = coluna;
            break;
          case 3:

            this.marcaDuploColuna3 = item;
            this.grafico3ColunaModel5 = null;
            this.grafico3ColunaModel5 = coluna;
            break;
          case 4:

            this.marcaDuploColuna4 = item;
            this.grafico3ColunaModel7 = null;
            this.grafico3ColunaModel7 = coluna;
            break;
          case 5:

            this.marcaDuploColuna5 = item;
            this.grafico3ColunaModel9 = null;
            this.grafico3ColunaModel9 = coluna;
            break;
        }

      }, (error) => { console.error(error) },
        () => { })


    var filtroColun2 = this.carregaFiltros();
    filtroColun2.Marca = new Array<PadraoComboFiltro>();
    filtroColun2.Marca.push(item);
    filtroColun2.Sequencia = nCol;
    filtroColun2.Onda = new Array<PadraoComboFiltro>();
    filtroColun2.Onda.push(onda2);
    this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColun2)
      .subscribe((coluna2: Array<GraficoImagemPosicionamento>) => {

        switch (nCol) {
          case 1:
            this.marcaDuploColuna1 = item;
            this.grafico3ColunaModel2 = null;
            this.grafico3ColunaModel2 = coluna2;
            break;
          case 2:

            this.marcaDuploColuna2 = item;
            this.grafico3ColunaModel4 = null;
            this.grafico3ColunaModel4 = coluna2;
            break;
          case 3:

            this.marcaDuploColuna3 = item;
            this.grafico3ColunaModel6 = null;
            this.grafico3ColunaModel6 = coluna2;
            break;
          case 4:

            this.marcaDuploColuna4 = item;
            this.grafico3ColunaModel8 = null;
            this.grafico3ColunaModel8 = coluna2;
            break;
          case 5:

            this.marcaDuploColuna5 = item;
            this.grafico3ColunaModel10 = null;
            this.grafico3ColunaModel10 = coluna2;
            break;
        }

      }, (error) => { console.error(error) },
        () => { })




      var filtroColun3 = this.carregaFiltros();
      filtroColun3.Marca = new Array<PadraoComboFiltro>();
      filtroColun3.Marca.push(item);
      filtroColun3.Sequencia = nCol;
      filtroColun3.Onda = new Array<PadraoComboFiltro>();
      filtroColun3.Onda.push(onda3);
      this.dashBoardFiveService.carregarGraficoComparativoMarcas2(filtroColun3)
        .subscribe((coluna3: Array<GraficoImagemPosicionamento>) => {

          switch (nCol) {
            case 1:
              this.marcaDuploColuna1 = item;
              this.grafico3ColunaModel1_3 = coluna3;
              break;
            case 2:

              this.marcaDuploColuna2 = item;
              this.grafico3ColunaModel3_3 = coluna3;
              break;
            case 3:

              this.marcaDuploColuna3 = item;
              this.grafico3ColunaModel5_3 = coluna3;
              break;
            case 4:

              this.marcaDuploColuna4 = item;
              this.grafico3ColunaModel7_3 = coluna3;
              break;
            case 5:

              this.marcaDuploColuna5 = item;
              this.grafico3ColunaModel9_3 = coluna3;
              break;
          }

     

        }, (error) => { console.error(error) },
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
      }, (error) => { console.error(error) });
  }

  // exportToPptComparativoExperiencia() {
  //   this.conversorPowerpointService.captureScreenPPTAlternative(
  //     'comparativo-experiencia',
  //     "NPS comparativo por Experiencia",
  //   );
  // }

  exportToPptComparativoMarcas() {

    var nome = this.translate.instant('dashboard-five-grafico1-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'comparativo-marcas',
      nome,
    );
  }


  exportToPptComparativoMarcasEvolutivo() {

    var nome = this.translate.instant('dashboard-five-grafico2-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'comparativo-marcas-evolutivo',
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

  exportToPptComparativoMarcasDuplo() {

    var nome = this.translate.instant('dashboard-five-grafico3-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'comparativo-marcas-duplo',
      nome,
    );
  }



  async exportToExcelGrafico1() {
    var filtroPadrao = this.carregaFiltrosExcel();
    var titulo = this.translate.instant('dashboard-five-grafico1-titulo');
    filtroPadrao.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardFiveComparativoMarcas(filtroPadrao)
      .subscribe((result) => {

        saveAs(
          result,
          titulo + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => { console.error(error) });
  }

  async exportToExcelGrafico2() {
    var filtroPadrao = this.carregaFiltrosExcel();
    var titulo = this.translate.instant('dashboard-five-grafico2-titulo');
    filtroPadrao.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardFiveEvolutivoMarcas(filtroPadrao)
      .subscribe((result) => {

        saveAs(
          result,
          titulo + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => { console.error(error) });
  }


  async exportToExcelGrafico3() {
    var filtroPadrao = this.carregaFiltrosExcelGraficoDuplo();
    var titulo = this.translate.instant('dashboard-five-grafico3-titulo');
    filtroPadrao.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardFiveComparativoMarcasDuplo(filtroPadrao)
      .subscribe((result) => {

        saveAs(
          result,
          titulo + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => { console.error(error) });
  }


 ajusteHeight(dados: Array<GraficoImagemPosicionamento>) {

    if (dados.length <= 0) {
      return "1000px";
    }
    else {
      var autoAjust = dados.length * 70;

      return autoAjust.toString() + "px";
    }
  }

  ajusteHeight2(dados: Array<GraficoImagemPosicionamento>) {

    if (dados.length <= 0) {
      return "1388px";
    }
    else {
      var autoAjust = dados.length * 66;

      autoAjust += 20;

      return autoAjust.toString() + "px";
    }
  }


  ajusteHeight1(dados: Array<GraficoImagemPosicionamento>) {

    if (dados.length <= 0) {
      return "1388px";
    }
    else {
      var autoAjust = dados.length * 66;

      autoAjust += 30;

      if (autoAjust < 500)
        return "500px";
      else
        return autoAjust.toString() + "px";
    }
  }



}
