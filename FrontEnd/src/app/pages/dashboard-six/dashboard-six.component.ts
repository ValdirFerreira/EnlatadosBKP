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
import { GraficoImagemPosicionamento } from 'src/app/models/grafico-Imagem-posicionamento/GraficoImagemPosicionamento';
import { DashBoardSixService } from 'src/app/services/dashboard-six.service';
import { GraficoBVCTop10 } from 'src/app/models/GraficoBVCTop10/GraficoBVCTop10';
import { GraficoBVCEvolutivo } from 'src/app/models/GraficoBVCEvolutivo/GraficoBVCEvolutivo';
import { LogService } from 'src/app/services/log.service';



@Component({
  selector: 'app-dashboard-six',
  templateUrl: './dashboard-six.component.html',
  styleUrls: ['./dashboard-six.component.scss']
})

export class DashboardSixComponent implements OnInit {


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
  grafico3ColunaModel3 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel4 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel5 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel6 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel7 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel8 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel9 = new Array<GraficoImagemPosicionamento>();
  grafico3ColunaModel10 = new Array<GraficoImagemPosicionamento>();

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

  ondaDuploColuna2: PadraoComboFiltro;
  ondaDuploColuna2_2: PadraoComboFiltro;

  ondaDuploColuna3: PadraoComboFiltro;
  ondaDuploColuna3_2: PadraoComboFiltro;

  ondaDuploColuna4: PadraoComboFiltro;
  ondaDuploColuna4_2: PadraoComboFiltro;

  ondaDuploColuna5: PadraoComboFiltro;
  ondaDuploColuna5_2: PadraoComboFiltro;

  // Filtros de Marca para utilização da geração do Excel Grafico Comparativo Marcas 
  marcaDuploColuna1: PadraoComboFiltro;
  marcaDuploColuna2: PadraoComboFiltro;
  marcaDuploColuna3: PadraoComboFiltro;
  marcaDuploColuna4: PadraoComboFiltro;
  marcaDuploColuna5: PadraoComboFiltro;
  // Filtros de Marca para utilização da geração do Excel Grafico Comparativo Marcas 

  // Filtros de Marca para utilização da geração do Excel Grafico Comparativo Marcas 

  graficoBVCTop10 = Array<GraficoBVCTop10>();

  graficoBVCEvolutivo1 = Array<GraficoBVCEvolutivo>()
  graficoBVCEvolutivo2 = Array<GraficoBVCEvolutivo>()
  graficoBVCEvolutivo3 = Array<GraficoBVCEvolutivo>()
  graficoBVCEvolutivo4 = Array<GraficoBVCEvolutivo>()

  ondaTopColuna1: PadraoComboFiltro;
  ondaTopColuna2: PadraoComboFiltro;
  ondaTopColuna3: PadraoComboFiltro;
  ondaTopColuna4: PadraoComboFiltro;

  ativaGraficoEvolutivoMarcas: boolean = false;
  ativaGraficoComparativoMarcasDuplo: boolean = false;

  ModelAtributo = new PadraoComboFiltro();

  paginaAtiva: boolean = true;


  constructor(public router: Router,
    public menuService: MenuService,
    public filtroService: FiltroGlobalService, private downloadArquivoService: DownloadArquivoService, private dashBoardSixService: DashBoardSixService,
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

    this.menuService.nomePage = this.translate.instant('navbar.dashboard-six');
    this.menuService.activeMenu = 6;
    this.menuService.menuSelecao = "6"

    EventEmitterService.get("emit-dashboard-six").subscribe((x) => {

      this.paginaAtiva = false;

      this.carregarGraficos();

      this.logService.GravaLogRota(this.router.url).subscribe(
      );


    })

  }


  carregarGraficos() {
    this.filtroService.FiltroMarcas(this.filtroService.ModelRegiao)
      .subscribe((response: Array<PadraoComboFiltro>) => {
        this.filtroService.listaMarcas = response;


        this.marcaEvolutivo = null;

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
        this.ondaDuploColuna2 = null;
        this.ondaDuploColuna2_2 = null;
        this.ondaDuploColuna3 = null;
        this.ondaDuploColuna3_2 = null;
        this.ondaDuploColuna4 = null;
        this.ondaDuploColuna4_2 = null;
        this.ondaDuploColuna5 = null;
        this.ondaDuploColuna5_2 = null;

        this.ondaTopColuna1 = null;
        this.ondaTopColuna2 = null;
        this.ondaTopColuna3 = null;
        this.ondaTopColuna4 = null;


        this.carregaFiltroMarcas();
        this.carregaFiltroOndas();

        this.filtroService.FiltroBVC()
          .subscribe((response: Array<PadraoComboFiltro>) => {

            this.ModelAtributo = response[0];

            this.carregarGraficoBVCTop10Marcas();
            this.carregarGraficoBVCEvolutivo();
            // this.carregarGraficoEvolutivoMarcasDuploFirstLoad();
            // this.carregarGraficoEvolutivoMarcasFirstLoad(this.marcaEvolutivo);

            this.paginaAtiva = true;

          }, (error) => console.error(error),
            () => {
            }
          )



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
    filtros.ParamTipo = 1;
    return filtros;
  }


  carregarGraficoBVCTop10Marcas() {

    var filtro = this.carregaFiltros();
    filtro.Sequencia = 100;
    this.dashBoardSixService.carregarGraficoBVCTop10Marcas(filtro)
      .subscribe((grafico: Array<GraficoBVCTop10>) => {
        this.graficoBVCTop10 = grafico
      }, (error) => { console.error(error) },
        () => { })
  }

  carregarGraficoBVCEvolutivo() {


    var filtro1 = this.carregaFiltros();
    filtro1.Sequencia = 100;
    filtro1.ParamTipo = this.ModelAtributo.IdItem == 0 ? 1 : this.ModelAtributo.IdItem;
    filtro1.Onda = new Array<PadraoComboFiltro>();
    filtro1.Onda.push(this.ondaTopColuna1);
    filtro1.ParamOndaAtual = this.filtroService.ModelOnda.IdItem;
    this.dashBoardSixService.CarregarGraficoBVCEvolutivo(filtro1)
      .subscribe((grafico: Array<GraficoBVCEvolutivo>) => {
        this.graficoBVCEvolutivo1 = grafico
      }, (error) => { console.error(error) },
        () => { })



    var filtro2 = this.carregaFiltros();
    filtro2.Sequencia = 71;
    filtro2.ParamTipo = this.ModelAtributo.IdItem == 0 ? 1 : this.ModelAtributo.IdItem;
    filtro2.Onda = new Array<PadraoComboFiltro>();
    filtro2.Onda.push(this.ondaTopColuna2);
    filtro2.ParamOndaAtual = this.filtroService.ModelOnda.IdItem;
    this.dashBoardSixService.CarregarGraficoBVCEvolutivo(filtro2)
      .subscribe((grafico: Array<GraficoBVCEvolutivo>) => {
        this.graficoBVCEvolutivo2 = grafico
      }, (error) => { console.error(error) },
        () => { })

    var filtro3 = this.carregaFiltros();
    filtro3.Sequencia = 72;
    filtro3.ParamTipo = this.ModelAtributo.IdItem == 0 ? 1 : this.ModelAtributo.IdItem;
    filtro3.Onda = new Array<PadraoComboFiltro>();
    filtro3.Onda.push(this.ondaTopColuna3);
    filtro3.ParamOndaAtual = this.filtroService.ModelOnda.IdItem;
    this.dashBoardSixService.CarregarGraficoBVCEvolutivo(filtro3)
      .subscribe((grafico: Array<GraficoBVCEvolutivo>) => {
        this.graficoBVCEvolutivo3 = grafico
      }, (error) => { console.error(error) },
        () => { })

    var filtro4 = this.carregaFiltros();
    filtro4.Sequencia = 73;
    filtro4.ParamTipo = this.ModelAtributo.IdItem == 0 ? 1 : this.ModelAtributo.IdItem;
    filtro4.Onda = new Array<PadraoComboFiltro>();
    filtro4.Onda.push(this.ondaTopColuna4);
    filtro4.ParamOndaAtual = this.filtroService.ModelOnda.IdItem;
    this.dashBoardSixService.CarregarGraficoBVCEvolutivo(filtro4)
      .subscribe((grafico: Array<GraficoBVCEvolutivo>) => {
        this.graficoBVCEvolutivo4 = grafico
      }, (error) => { console.error(error) },
        () => { })

  }

  onchangeOndaColuna1(item: PadraoComboFiltro) {
    this.carregarGraficoBVCTop10Marcas();
    this.carregarGraficoBVCEvolutivo();
  }

  carregarGraficoEvolutivoMarcasFirstLoad(item: PadraoComboFiltro = null) {
    var marca = item == null ? this.filtroService.listaMarcas[0] : item;

    var filtroColuna1 = this.carregaFiltros();
    filtroColuna1.Sequencia = 60;
    filtroColuna1.Marca = new Array<PadraoComboFiltro>();
    filtroColuna1.Marca.push(marca);
    filtroColuna1.Onda = new Array<PadraoComboFiltro>();
    filtroColuna1.Onda.push(this.ondaColuna1);
    this.marcaEvolutivo = marca;


    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna1)
      .subscribe((coluna1: Array<GraficoImagemPosicionamento>) => {

        this.grafico2ColunaModel1 = coluna1;

      }, (error) => console.error(error),
        () => { })

    var filtroColuna2 = this.carregaFiltros();
    filtroColuna2.Sequencia = 61;
    filtroColuna2.Marca = new Array<PadraoComboFiltro>();
    filtroColuna2.Marca.push(marca);
    filtroColuna2.Onda = new Array<PadraoComboFiltro>();
    filtroColuna2.Onda.push(this.ondaColuna2);

    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna2)
      .subscribe((coluna2: Array<GraficoImagemPosicionamento>) => {
        this.grafico2ColunaModel2 = coluna2;


      }, (error) => console.error(error),
        () => { })

    var filtroColuna3 = this.carregaFiltros();
    filtroColuna3.Sequencia = 62;
    filtroColuna3.Marca = new Array<PadraoComboFiltro>();
    filtroColuna3.Marca.push(marca);
    filtroColuna3.Onda = new Array<PadraoComboFiltro>();
    filtroColuna3.Onda.push(this.ondaColuna3);

    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna3)
      .subscribe((coluna3: Array<GraficoImagemPosicionamento>) => {
        this.grafico2ColunaModel3 = coluna3

      }, (error) => console.error(error),
        () => { })

    var filtroColuna4 = this.carregaFiltros();
    filtroColuna4.Sequencia = 63;
    filtroColuna4.Marca = new Array<PadraoComboFiltro>();
    filtroColuna4.Marca.push(marca);
    filtroColuna4.Onda = new Array<PadraoComboFiltro>();
    filtroColuna4.Onda.push(this.ondaColuna4);

    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna4)
      .subscribe((coluna4: Array<GraficoImagemPosicionamento>) => {
        this.grafico2ColunaModel4 = coluna4

        this.ativaGraficoEvolutivoMarcas = true;

      }, (error) => console.error(error),
        () => { })


  }


  carregarGraficoEvolutivoMarcasDuploFirstLoad() {


    var filtroColuna1 = this.carregaFiltros();
    filtroColuna1.Sequencia = 40;
    filtroColuna1.Marca = new Array<PadraoComboFiltro>();
    filtroColuna1.Marca.push(this.marcaDuploColuna1);
    filtroColuna1.Onda = new Array<PadraoComboFiltro>();
    filtroColuna1.Onda.push(this.ondaDuploColuna1);

    // Grafico 1
    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna1)
      .subscribe((colunaPos: Array<GraficoImagemPosicionamento>) => {

        if (this.grafico3ColunaModel1)
          this.grafico3ColunaModel1 = colunaPos;
        else {
          this.grafico3ColunaModel1 = null;
        }



      }, (error) => console.error(error),
        () => { })

    //Grafico 1
    var filtroColuna2 = this.carregaFiltros();
    filtroColuna2.Sequencia = 8;
    filtroColuna2.Marca = new Array<PadraoComboFiltro>();
    filtroColuna2.Marca.push(this.marcaDuploColuna1);
    filtroColuna2.Onda = new Array<PadraoComboFiltro>();
    filtroColuna2.Onda.push(this.ondaDuploColuna1_2);

    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna2)
      .subscribe((coluna2: Array<GraficoImagemPosicionamento>) => {
        this.grafico3ColunaModel2 = coluna2

      }, (error) => console.error(error),
        () => { })


    // Grafico 2
    var filtroColuna3 = this.carregaFiltros();
    filtroColuna3.Sequencia = 9;
    filtroColuna3.Marca = new Array<PadraoComboFiltro>();
    filtroColuna3.Marca.push(this.marcaDuploColuna2);
    filtroColuna3.Onda = new Array<PadraoComboFiltro>();
    filtroColuna3.Onda.push(this.ondaDuploColuna2);

    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna3)
      .subscribe((coluna3: Array<GraficoImagemPosicionamento>) => {

        if (this.grafico3ColunaModel3)
          this.grafico3ColunaModel3 = coluna3
        else this.grafico3ColunaModel3 = null;

      }, (error) => console.error(error),
        () => { })

    // Grafico 2
    var filtroColuna4 = this.carregaFiltros();
    filtroColuna4.Sequencia = 10;
    filtroColuna4.Marca = new Array<PadraoComboFiltro>();
    filtroColuna4.Marca.push(this.marcaDuploColuna2);
    filtroColuna4.Onda = new Array<PadraoComboFiltro>();
    filtroColuna4.Onda.push(this.ondaDuploColuna2_2);

    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna4)
      .subscribe((coluna4: Array<GraficoImagemPosicionamento>) => {

        if (this.grafico3ColunaModel4)
          this.grafico3ColunaModel4 = coluna4
        else
          this.grafico3ColunaModel4 = null;

      }, (error) => console.error(error),
        () => { })

    // Grafico 3
    var filtroColuna5 = this.carregaFiltros();
    filtroColuna5.Sequencia = 11;
    filtroColuna5.Marca = new Array<PadraoComboFiltro>();
    filtroColuna5.Marca.push(this.marcaDuploColuna3);
    filtroColuna5.Onda = new Array<PadraoComboFiltro>();
    filtroColuna5.Onda.push(this.ondaDuploColuna3);

    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna5)
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

    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna6)
      .subscribe((coluna6: Array<GraficoImagemPosicionamento>) => {
        this.grafico3ColunaModel6 = coluna6

      }, (error) => console.error(error),
        () => { })


    // Grafico 4
    var filtroColuna7 = this.carregaFiltros();
    filtroColuna7.Sequencia = 13;
    filtroColuna7.Marca = new Array<PadraoComboFiltro>();
    filtroColuna7.Marca.push(this.marcaDuploColuna4);
    filtroColuna7.Onda = new Array<PadraoComboFiltro>();
    filtroColuna7.Onda.push(this.ondaDuploColuna4);

    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna7)
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

    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna8)
      .subscribe((coluna8: Array<GraficoImagemPosicionamento>) => {
        this.grafico3ColunaModel8 = coluna8


      }, (error) => console.error(error),
        () => { })

    // Grafico 5
    var filtroColuna9 = this.carregaFiltros();
    filtroColuna9.Sequencia = 15;
    filtroColuna9.Marca = new Array<PadraoComboFiltro>();
    filtroColuna9.Marca.push(this.marcaDuploColuna5);
    filtroColuna9.Onda = new Array<PadraoComboFiltro>();
    filtroColuna9.Onda.push(this.ondaDuploColuna5);

    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna9)
      .subscribe((coluna9: Array<GraficoImagemPosicionamento>) => {
        this.grafico3ColunaModel9 = coluna9

      }, (error) => console.error(error),
        () => { })


    var filtroColuna10 = this.carregaFiltros();
    filtroColuna10.Sequencia = 16;
    filtroColuna10.Marca = new Array<PadraoComboFiltro>();
    filtroColuna10.Marca.push(this.marcaDuploColuna5);
    filtroColuna10.Onda = new Array<PadraoComboFiltro>();
    filtroColuna10.Onda.push(this.ondaDuploColuna5_2);

    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna10)
      .subscribe((coluna10: Array<GraficoImagemPosicionamento>) => {
        this.grafico3ColunaModel10 = coluna10

        this.ativaGraficoComparativoMarcasDuplo = true;

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


    if (!this.ondaTopColuna1) {
      this.ondaTopColuna1 = this.filtroService.listaOnda[3];
    }

    if (!this.ondaTopColuna2) {
      this.ondaTopColuna2 = this.filtroService.listaOnda[2];
    }

    if (!this.ondaTopColuna3) {
      this.ondaTopColuna3 = this.filtroService.listaOnda[1];
    }

    if (!this.ondaTopColuna4) {
      this.ondaTopColuna4 = this.filtroService.listaOnda[0];
    }



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
      this.ondaDuploColuna1 = this.filtroService.listaOnda[1];
    }
    if (!this.ondaDuploColuna1_2) {
      this.ondaDuploColuna1_2 = this.filtroService.listaOnda[0];
    }

    if (!this.ondaDuploColuna2) {
      this.ondaDuploColuna2 = this.filtroService.listaOnda[1];
    }
    if (!this.ondaDuploColuna2_2) {
      this.ondaDuploColuna2_2 = this.filtroService.listaOnda[0];
    }


    if (!this.ondaDuploColuna3) {
      this.ondaDuploColuna3 = this.filtroService.listaOnda[1];
    }
    if (!this.ondaDuploColuna3_2) {
      this.ondaDuploColuna3_2 = this.filtroService.listaOnda[0];
    }


    if (!this.ondaDuploColuna4) {
      this.ondaDuploColuna4 = this.filtroService.listaOnda[1];
    }
    if (!this.ondaDuploColuna4_2) {
      this.ondaDuploColuna4_2 = this.filtroService.listaOnda[0];
    }


    if (!this.ondaDuploColuna5) {
      this.ondaDuploColuna5 = this.filtroService.listaOnda[1];
    }
    if (!this.ondaDuploColuna5_2) {
      this.ondaDuploColuna5_2 = this.filtroService.listaOnda[0];
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

    filtros.OndaDuploColuna2 = this.ondaDuploColuna2;
    filtros.OndaDuploColuna2_2 = this.ondaDuploColuna2_2;

    filtros.OndaDuploColuna3 = this.ondaDuploColuna3;
    filtros.OndaDuploColuna3_2 = this.ondaDuploColuna3_2;

    filtros.OndaDuploColuna4 = this.ondaDuploColuna4;
    filtros.OndaDuploColuna4_2 = this.ondaDuploColuna4_2;

    filtros.OndaDuploColuna5 = this.ondaDuploColuna5;
    filtros.OndaDuploColuna5_2 = this.ondaDuploColuna5_2;

    filtros.MarcaDuploColuna1 = this.marcaDuploColuna1;
    filtros.MarcaDuploColuna2 = this.marcaDuploColuna2;
    filtros.MarcaDuploColuna3 = this.marcaDuploColuna3;
    filtros.MarcaDuploColuna4 = this.marcaDuploColuna4;
    filtros.MarcaDuploColuna5 = this.marcaDuploColuna5;

    filtros.CodUser = this.session.getCodUserSession();
    filtros.CodIdioma = 1;// this.authService.idDefaultLangUser;
    return filtros;
  }

  public carregaFiltrosExcelTop() {

    var filtros = new FiltroPadraoExcel();
    var list = [];
    list.push(this.filtroService.ModelTarget);
    filtros.Target = list;
    filtros.Regiao = this.filtroService.ModelRegiao;
    filtros.Demografico = this.filtroService.ModelDemografico;
    filtros.Onda.push(this.filtroService.ModelOnda);
    filtros.ParamTipo = this.ModelAtributo.IdItem == 0 ? 1 : this.ModelAtributo.IdItem;

    filtros.Onda1.push(this.ondaTopColuna1);
    filtros.Onda2.push(this.ondaTopColuna2);
    filtros.Onda3.push(this.ondaTopColuna3);
    filtros.Onda4.push(this.ondaTopColuna4);

    filtros.ParamOndaAtual = this.filtroService.ModelOnda.IdItem;


    filtros.CodUser = this.session.getCodUserSession();
    filtros.CodIdioma = 1;// this.authService.idDefaultLangUser;
    return filtros;
  }

  onChangeAtributo(atributo: PadraoComboFiltro) {
    this.ModelAtributo = atributo;
    this.carregarGraficoBVCEvolutivo();
  }

  // grafico Comparativo de marcas
  onchangeMarca(item: PadraoComboFiltro, nCol: number) {

    var filtroColuna = this.carregaFiltros();
    filtroColuna.Marca = new Array<PadraoComboFiltro>();
    filtroColuna.Marca.push(item);
    filtroColuna.Sequencia = nCol;
    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna)
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

    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna1)
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


  onchangeOndaTop(item: PadraoComboFiltro, nCol: number) {

    var filtro = this.carregaFiltros();
    filtro.Sequencia = 1;
    filtro.Onda = new Array<PadraoComboFiltro>();
    filtro.Onda.push(item);
    filtro.Sequencia = nCol;
    filtro.ParamTipo = this.ModelAtributo.IdItem == 0 ? 1 : this.ModelAtributo.IdItem;
    filtro.ParamOndaAtual = this.filtroService.ModelOnda.IdItem;
    this.dashBoardSixService.CarregarGraficoBVCEvolutivo(filtro)
      .subscribe((grafico: Array<GraficoBVCEvolutivo>) => {

        switch (nCol) {
          case 1:
            this.graficoBVCEvolutivo1 = grafico
            this.ondaTopColuna1 = item;
            break;
          case 2:
            this.graficoBVCEvolutivo2 = grafico
            this.ondaTopColuna2 = item;
            break;
          case 3:
            this.graficoBVCEvolutivo3 = grafico
            this.ondaTopColuna3 = item;
            break;
          case 4:
            this.graficoBVCEvolutivo4 = grafico
            this.ondaTopColuna4 = item;
            break;
            // case 5:
            //   this.grafico2ColunaModel5 = coluna
            break;
        }

      }, (error) => { console.error(error) },
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
      case 2:
        filtroColuna.Marca.push(this.marcaDuploColuna1);
        break;

      case 3:
      case 4:
        filtroColuna.Marca.push(this.marcaDuploColuna2);
        break;

      case 5:
      case 6:
        filtroColuna.Marca.push(this.marcaDuploColuna3);
        break;

      case 7:
      case 8:
        filtroColuna.Marca.push(this.marcaDuploColuna4);
        break;

      case 9:
      case 10:
        filtroColuna.Marca.push(this.marcaDuploColuna5);
        break;

    }


    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna)
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
          case 3:
            this.grafico3ColunaModel3 = coluna;
            this.ondaDuploColuna2 = item;
            break;
          case 4:
            this.grafico3ColunaModel4 = coluna;
            this.ondaDuploColuna2_2 = item;
            break;

          case 5:
            this.grafico3ColunaModel5 = coluna;
            this.ondaDuploColuna3 = item;
            break;

          case 6:
            this.grafico3ColunaModel6 = coluna;
            this.ondaDuploColuna3_2 = item;
            break;

          case 7:
            this.grafico3ColunaModel7 = coluna;
            this.ondaDuploColuna4 = item;
            break;

          case 8:
            this.grafico3ColunaModel8 = coluna;
            this.ondaDuploColuna4_2 = item;
            break;

          case 9:
            this.grafico3ColunaModel9 = coluna;
            this.ondaDuploColuna5 = item;
            break;

          case 10:
            this.grafico3ColunaModel10 = coluna;
            this.ondaDuploColuna5_2 = item;
            break;
            // case 5:
            //   this.grafico2ColunaModel5 = coluna
            break;
        }

      }, (error) => { console.error(error) },
        () => { })


  }

  onchangeDuploMarca(item: PadraoComboFiltro, nCol: number) {

    var onda1 = new PadraoComboFiltro();
    var onda2 = new PadraoComboFiltro();

    switch (nCol) {
      case 1:

        onda1 = this.ondaDuploColuna1;
        onda2 = this.ondaDuploColuna1_2;
        break;

      case 2:
        onda1 = this.ondaDuploColuna2;
        onda2 = this.ondaDuploColuna2_2;
        break;

      case 3:
        onda1 = this.ondaDuploColuna3;
        onda2 = this.ondaDuploColuna3_2;
        break;

      case 4:
        onda1 = this.ondaDuploColuna4;
        onda2 = this.ondaDuploColuna4_2;
        break;

      case 5:
        onda1 = this.ondaDuploColuna5;
        onda2 = this.ondaDuploColuna5_2;
        break;
    }


    var filtroColuna = this.carregaFiltros();
    filtroColuna.Marca = new Array<PadraoComboFiltro>();
    filtroColuna.Marca.push(item);
    filtroColuna.Sequencia = nCol;
    filtroColuna.Onda = new Array<PadraoComboFiltro>();
    filtroColuna.Onda.push(onda1);
    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColuna)
      .subscribe((coluna: Array<GraficoImagemPosicionamento>) => {

        switch (nCol) {
          case 1:
            this.marcaDuploColuna1 = item;
            this.grafico3ColunaModel1 = coluna;
            break;
          case 2:

            this.marcaDuploColuna2 = item;
            this.grafico3ColunaModel3 = coluna;
            break;
          case 3:

            this.marcaDuploColuna3 = item;
            this.grafico3ColunaModel5 = coluna;
            break;
          case 4:

            this.marcaDuploColuna4 = item;
            this.grafico3ColunaModel7 = coluna;
            break;
          case 5:

            this.marcaDuploColuna5 = item;
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
    this.dashBoardSixService.carregarGraficoComparativoMarcas(filtroColun2)
      .subscribe((coluna2: Array<GraficoImagemPosicionamento>) => {

        switch (nCol) {
          case 1:
            this.marcaDuploColuna1 = item;
            this.grafico3ColunaModel2 = coluna2;
            break;
          case 2:

            this.marcaDuploColuna2 = item;
            this.grafico3ColunaModel4 = coluna2;
            break;
          case 3:

            this.marcaDuploColuna3 = item;
            this.grafico3ColunaModel6 = coluna2;
            break;
          case 4:

            this.marcaDuploColuna4 = item;
            this.grafico3ColunaModel8 = coluna2;
            break;
          case 5:

            this.marcaDuploColuna5 = item;
            this.grafico3ColunaModel10 = coluna2;
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

    var nome = this.translate.instant('dashboard-six-grafico1-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'comparativo-marcas',
      nome,
    );
  }

  exportToPptEvolutivoMarcas() {
    var nome = this.translate.instant('dashboard-six-grafico2-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'comparativo-marcas-evolutivo',
      nome
    );
  }

  exportToPptComparativoMarcasDuplo() {

    var nome = this.translate.instant('dashboard-six-grafico3-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'comparativo-marcas-duplo-mercado',
      nome,
    );
  }

  exportToPptComparativoMarcasEvolutivo() {

    var nome = this.translate.instant('dashboard-six-grafico4-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'comparativo-marcas-mercado-onda',
      nome,
    );
  }





  async exportToExcelGrafico1() {
    var filtroPadrao = this.carregaFiltros();
    var titulo = this.translate.instant('dashboard-six-grafico1-titulo');
    filtroPadrao.TituloGrafico = titulo;

    await this.downloadArquivoService.
      DownloadDashboardSixGraficoBVCTop10Marcas(filtroPadrao)
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
    var filtroPadrao = this.carregaFiltrosExcelTop();
    var titulo = this.translate.instant('dashboard-six-grafico2-titulo');
    filtroPadrao.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardSixGraficoBVCEvolutivo(filtroPadrao)
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
    var titulo = this.translate.instant('dashboard-six-grafico3-titulo');
    filtroPadrao.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardSixComparativoMarcasDuplo(filtroPadrao)
      .subscribe((result) => {

        saveAs(
          result,
          titulo + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => { console.error(error) });
  }

  async exportToExcelGrafico4() {
    var filtroPadrao = this.carregaFiltrosExcel();
    var titulo = this.translate.instant('dashboard-six-grafico4-titulo');
    filtroPadrao.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardSixEvolutivoMarcas(filtroPadrao)
      .subscribe((result) => {

        saveAs(
          result,
          titulo + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => { console.error(error) });
  }






}
