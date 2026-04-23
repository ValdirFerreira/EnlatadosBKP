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
import { GraficoLinhasModel } from 'src/app/models/grafico-highchart/grafico-highchart';

import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';
import { Session } from '../home/guards/session';
import { LogService } from 'src/app/services/log.service';
import { DashBoardNineService } from 'src/app/services/dashboard-nine-service';
import { TabelaAdHocAtributo, TabelaPadraoAdHoc } from 'src/app/models/tabela-padrao/TabelaPadraoAdHoc';



@Component({
  selector: 'app-dashboard-nine',
  templateUrl: './dashboard-nine.component.html',
  styleUrls: ['./dashboard-nine.component.scss']
})
export class DashboardNineComponent implements OnInit {

  firstLoad: boolean = true;

  ativaComparativoMarcas: boolean = true;
  graficoComparativoMarcasModel = new GraficoLinhasModel();

  graficoImagemPuraEvolutivoColuna1Model = new GraficoLinhasModel();
  graficoImagemPuraEvolutivoColuna2Model = new GraficoLinhasModel();

  check1: boolean = true;

  // Filtros de Marca para utilização da geração do Excel Grafico Comparativo Marcas 
  marcaColuna1: PadraoComboFiltro;
  marcaColuna2: PadraoComboFiltro;
  marcaColuna3: PadraoComboFiltro;
  marcaColuna4: PadraoComboFiltro;
  marcaColuna5: PadraoComboFiltro;

  marcasSelecionadasGrafico1 = new Array<PadraoComboFiltro>();
  marcasSelecionadasGrafico3 = new Array<PadraoComboFiltro>();
  ModelAtributo = new PadraoComboFiltro();
  // Filtros de Marca para utilização da geração do Excel Grafico Comparativo Marcas 

  descBaseComparativoMarcas: string = "";
  descVerificaBaseGraficoLinha: string = "";
  descVerificaBaseEvolutivoMarcas: string = "";
  descVerificaBaseEvolutivoMarcas2: string = "";


  ativaEvolutivoMarcasLinhas: boolean = true;
  ativaEvolutivoMarcasLinhas2: boolean = true;
  ativaTabelaAdHoc: boolean = true;

  ativaTabelaAdHocBloco6: boolean = true;
  tabelaPadraoAdHocBloco6 = new TabelaPadraoAdHoc();


  ativaTabelaAdHocBloco10: boolean = true;
  tabelaPadraoAdHocBloco10 = new TabelaPadraoAdHoc();
  marcatabelaAdHocBloco10: PadraoComboFiltro;

  ativaTabelaAdHocBloco2: boolean = true;
  tabelaPadraoAdHocBloco2 = new TabelaPadraoAdHoc();

  paginaAtiva: boolean = true;

  tabelaPadraoAdHoc = new TabelaPadraoAdHoc();
  marcatabelaAdHoc: PadraoComboFiltro;


  marcatabelaAdHocBloco6: PadraoComboFiltro;

  listaMarcas10: Array<PadraoComboFiltro>;




  constructor(public router: Router,
    public menuService: MenuService,
    public filtroService: FiltroGlobalService, private downloadArquivoService: DownloadArquivoService,
    private translate: TranslateService, private conversorPowerpointService: ConversorPowerpointService,
    private dashBoardNineService: DashBoardNineService,
    private session: Session,
    private logService: LogService,
  ) { }

  ngOnInit(): void {

    window.scroll({
      top: 0,
      left: 0,
      behavior: 'smooth'
    });

    this.menuService.nomePage = this.translate.instant('navbar.dashboard-nine');
    this.ModelAtributo.IdItem = 1;

    this.menuService.activeMenu = 4;
    this.menuService.menuSelecao = "4"


    EventEmitterService.get("emit-dashboard-nine").subscribe((x) => {
      this.menuService.nomePage = this.translate.instant('navbar.dashboard-nine');
      this.paginaAtiva = false;
      this.carregarGraficos();

      this.logService.GravaLogRota(this.router.url).subscribe(
      );

    })

    // this.carregarEvolutivoMarcas();

  }


  FiltroAdHoc() {
    this.filtroService.FiltroAdHoc()
      .subscribe((response: Array<PadraoComboFiltro>) => {
        this.listaMarcas10 = response;

      }, (error) => console.error(error),
        () => {
        }
      )
  }



  onChangeAtributo(atributo: PadraoComboFiltro) {

    this.ModelAtributo = atributo;
    this.carregarEvolutivoMarcas();
  }



  carregarGraficos() {

    this.marcaColuna1 = null;
    this.marcaColuna2 = null;
    this.marcaColuna3 = null;
    this.marcaColuna4 = null;
    this.marcaColuna5 = null;

    this.filtroService.FiltroAdHoc()
      .subscribe((response: Array<PadraoComboFiltro>) => {
        this.ModelAtributo = response[0];

      }, (error) => console.error(error),
        () => {
        }
      )


    this.filtroService.FiltroMarcasAdHoc(10)
      .subscribe((response: Array<PadraoComboFiltro>) => {
        this.marcatabelaAdHocBloco10 = response[0]
        this.listaMarcas10 = response;


        this.filtroService.FiltroMarcas(this.filtroService.ModelRegiao)
          .subscribe((response: Array<PadraoComboFiltro>) => {

            debugger
            this.filtroService.listaMarcas = response;
            // this.filtroService.ModelDemografico = response[0];

            this.marcaColuna1 = this.filtroService.listaMarcas[0]
            this.marcaColuna2 = this.filtroService.listaMarcas[1]
            this.marcaColuna3 = this.filtroService.listaMarcas[2]
            this.marcaColuna4 = this.filtroService.listaMarcas[3]
            this.marcaColuna5 = this.filtroService.listaMarcas[4]

            this.marcatabelaAdHocBloco6 = this.filtroService.listaMarcas[0]
            this.marcatabelaAdHoc = this.filtroService.listaMarcas[0]

            this.paginaAtiva = true;

            // this.carregaFiltroMarcas();
            // this.carregarEvolutivoMarcas();
            // this.TabelaAdHocAtributoBloco6();
            this.TabelaAdHocAtributoBloco10();
            // this.TabelaAdHocAtributoBloco2();
            this.filtroService.FiltroMarcasAdHoc(12)
              .subscribe((response: Array<PadraoComboFiltro>) => {
                this.listaMarcas = response;
                this.ModelMarcas = response[0]

                this.listaMarcas2 = response;
                this.ModelMarcas2 = response[0]

                this.montaImageMarca();
                //  this.carregarEvolutivoMarcas2();

                this.carregarEvolutivoMarcas();

                // this.montaImageMarca2();
                // this.carregarEvolutivoMarcas02();
              }, (error) => console.error(error),
                () => {
                }
              )


            // this.carregarGraficosFirstLoad();
            // this.carregarEvolutivoMarcas2();

            // this.filtroService.FiltroMarcas(this.filtroService.ModelRegiao)
            //   .subscribe((response: Array<PadraoComboFiltro>) => {
            //     // Neste momento somente devemos trazer    Ninho//Itambé//Piracanjuba
            //    // this.filtroService.listaMarcas = response.filter(x => x.IdItem == 13001 || x.IdItem == 13003 || x.IdItem == 13013);;
            //     this.marcatabelaAdHoc = this.filtroService.listaMarcas[0]
            //     this.carregarTabelaAdHocAtributo();

            //   }, (error) => console.error(error),
            //     () => {
            //     }
            //   )

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
    filtros.Onda = new Array<PadraoComboFiltro>();
    filtros.Onda.push(this.filtroService.ModelOnda);
    filtros.CodUser = this.session.getCodUserSession();
    filtros.CodIdioma = 1;// this.authService.idDefaultLangUser;
    return filtros;
  }


  changetipoBia(nGrafico: number) {

    switch (nGrafico) {
      case 1:

        break;

      case 2:
        this.carregarGraficosFirstLoad();
        break;

      case 3:
        this.ModelAtributo.IdItem = 1;
        this.ativaEvolutivoMarcasLinhas = false;
        this.carregarEvolutivoMarcas();
        break;

      default:
        break;
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

    filtros.Marca1.push(!this.marcaColuna1 ? this.filtroService.listaMarcas[0] : this.marcaColuna1)
    filtros.Marca2.push(!this.marcaColuna2 ? this.filtroService.listaMarcas[1] : this.marcaColuna2);
    filtros.Marca3.push(!this.marcaColuna3 ? this.filtroService.listaMarcas[2] : this.marcaColuna3);
    filtros.Marca4.push(!this.marcaColuna4 ? this.filtroService.listaMarcas[3] : this.marcaColuna4);
    filtros.Marca5.push(!this.marcaColuna5 ? this.filtroService.listaMarcas[4] : this.marcaColuna5);

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



  getArrayColors(tam: number): string[] {
    // let arrColors = ['#e87943', '#4fe605', '#cac196', '#c996d3', '#f892b1', '#18b5ad', '#fa2c15', '#2c3fdd', '#72061e', '#8aaab3', '#fef1c1', '#5dced2', '#022279', '#00ae5a', '#c677fc', '#d101d1', '#a81c26', '#194920', '#4c9d6d', '#8c6a23', '#13f1f7', '#afea45', '#f81d6e', '#c8ff0b', '#d2ffdf', '#fbb15c'];

    let arrColors = ['#F2994A', '#2D9CDB', '#cac196', '#c996d3', '#f892b1', '#18b5ad', '#fa2c15', '#2c3fdd', '#72061e', '#8aaab3', '#fef1c1', '#5dced2', '#022279', '#00ae5a', '#c677fc', '#d101d1', '#a81c26', '#194920', '#4c9d6d', '#8c6a23', '#13f1f7', '#afea45', '#f81d6e', '#c8ff0b', '#d2ffdf', '#fbb15c'];

    // while (tam > arrColors.length) {
    //   let newColor = this.generateColor();

    //   if (!arrColors.includes(newColor))
    //     arrColors.push(newColor);
    // }

    return arrColors;
  }

  verificaBaseComparativoMarcas() {
    this.descBaseComparativoMarcas = "";
    if (this.graficoComparativoMarcasModel) {
      this.graficoComparativoMarcasModel.Grafico.forEach(item => {
        item.data.forEach(obj => {
          if (obj.baseminima != "") {
            this.descBaseComparativoMarcas = obj.baseminima;
          }
        });
      });
    }
  }

  verificaBaseGraficoLinha(grafico: GraficoLinhasModel) {
    if (grafico) {
      grafico.Grafico.forEach(item => {
        item.data.forEach(obj => {
          if (obj.baseminima != "") {
            this.descVerificaBaseGraficoLinha = obj.baseminima;
          }
        });
      });
    }
  }

  carregarGraficosFirstLoad() {
    this.descVerificaBaseGraficoLinha = "";
    this.graficoLinha1();
    this.graficoLinha("container-coluna-1", 1, this.marcaColuna1);
    this.graficoLinha("container-coluna-2", 2, this.marcaColuna2);
    this.graficoLinha("container-coluna-3", 3, this.marcaColuna3);
    this.graficoLinha("container-coluna-4", 4, this.marcaColuna4);
    this.graficoLinha("container-coluna-5", 5, this.marcaColuna5);

  }


  graficoLinha1() {

    var graficoImagemPuraEvolutivoColunaModel = new GraficoLinhasModel();

    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Marca.push(this.filtroService.listaMarcas[0]);
    filtros.Sequencia = 0;

    this.dashBoardNineService.ImagemEvolutiva(filtros)
      .subscribe((response: GraficoLinhasModel) => {

        this.graficoImagemPuraEvolutivoColuna1Model = response;

      }, (error) => console.error(error),
        () => {

          var grafico = [] as any;
          var index = 0;
          // var cores = this.getArrayColors(graficoImagemPuraEvolutivoColunaModel.Grafico.length);
          var cores = ['#ffffff'];//this.graficoImagemPuraEvolutivoColuna1Model.Cores;

          // COR CINZA #75787c

          if (this.graficoImagemPuraEvolutivoColuna1Model?.Grafico?.length) {

            var lista = this.graficoImagemPuraEvolutivoColuna1Model.Grafico.reverse();

            lista?.forEach(((item) => {

              var objeto = {
                id: item.name,
                linkedTo: "legenda" + index,
                type: 'line',
                name: item.name,
                data: item.data,
                showInLegend: false,
                marker: {
                  symbol: "circle",
                  // fillColor: '#FFFFFF', // BOLINHA VAZADA
                  lineWidth: 2,
                  radius: 4,
                  lineColor: cores[index]
                }
              }

              var legenda = {
                id: "legenda" + index,
                name: item.name,
                data: [],
                color: cores[index],
                type: 'column',
                marker: { symbol: "circle", height: 10 }
              }

              grafico.push(objeto);
              grafico.push(legenda);
              index++;
            }));

            this.graficoImagemPuraEvolutivoColuna1Model.Grafico = grafico;

            var graficoImagemPuraEvolutivoColuna1ModelHighchart: any = {
              chart: {
                renderTo: 'container-coluna-0',
                inverted: true,
                spacingTop: 0,
                spacingLeft: 20,
                height: 400,
                width: 300,
                title: {
                  text: "Texto",
                  enabled: false,
                },
                labels: {
                  enabled: true,
                },
                type: "spline",
                // //backgroundColor: "#F8F8F8",
                style: {
                  fontFamily: "Poppins",
                },
              },
              title: {
                text: "",
                enabled: false,
              },
              legend: {
                align: "left",
                verticalAlign: "top",
                enabled: false,
                alignColumns: true,
                itemDistance: 30,
                spacingBottom: 0,

                itemStyle: {
                  fontSize: '14px',
                  fontWeight: 'normal',
                  fontFamily: 'Poppins',
                  fontStyle: 'normal',
                  color: '#585656',
                  lineHeight: '20px',
                  textAlign: 'left',

                },

              },
              xAxis: [{
                lineWidth: 1,
                gridLineColor: '#ffffff',
                lineColor: '#ffffff',
                gridLineWidth: 1,
                tickInterval: 1,
                // opposite: true,
                // reversed: false,
                // reversedStacks: false,
                categories: this.graficoImagemPuraEvolutivoColuna1Model.Periodos,
                labels: {

                  enabled: true,
                  floating: true,
                  //  rotation: 0,
                  align: 'left',
                  style: {

                    fontSize: '14px',
                    fontWeight: 'normal',
                    fontFamily: 'Poppins',
                    fontStyle: 'normal',
                    color: '#585656',
                    lineHeight: '20px',
                    textAlign: 'right',
                    width: 390,
                  },
                },


              }, {
                // top: '50%',
                // height: '50%',

                // gridLineColor: '#ffffff',
                lineColor: '#ffffff',
                gridLineWidth: 1,
                reversed: false,
                reversedStacks: true,
                tickInterval: 1,
                categories: this.graficoImagemPuraEvolutivoColuna1Model.Periodos,
                labels: {
                  enabled: true,
                  floating: true,
                  rotation: 90,
                  align: 'center',
                  style: {
                    fontSize: '14px',
                    fontWeight: 'normal',
                    fontFamily: 'Poppins',
                    fontStyle: 'normal',
                    color: '#585656',
                    lineHeight: '20px',
                    textAlign: 'left',
                    width: 700,
                  },
                },
              }],
              yAxis: {
                lineWidth: 1,
                gridLineColor: '#ffffff',
                lineColor: '#ffffff',

                gridLineWidth: 1,
                tickInterval: 1,
                label: {
                  enabled: true,

                },
                title: {
                  text: "undef",
                  enabled: false,
                },
                labels: {
                  enabled: false,
                  align: 'right',
                  // x: -2,

                  style: {
                    fontSize: '12px',
                    fontWeight: 'normal',
                    fontFamily: 'Poppins',
                    fontStyle: 'normal',
                    color: '#585656',
                  },
                },
              },

              tooltip: {
                shared: false,
                valueSuffix: "",
                enabled: false,
                useHTML: true,

                headerFormat: '',

                formatter(this: Highcharts.TooltipFormatterContextObject) {
                  var pointer = this.point as any;
                  return '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> Média  </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.y + '%</div>'
                    + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px">  Base </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.valorbase + '</div>  '
                  // + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> Período</div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.periodo + '</div> ';

                },

                // pointFormat:

                //   this.textoPoupUpGraficoLinha,

                // outside: true,
                // backgroundColor: "rgba(246, 246, 246, 1)",
                // borderRadius: 30,
                // borderColor: "#bbbbbb",
                // borderWidth: 1.5,
                // style: { opacity: 1, background: "rgba(246, 246, 246, 1)" },
                // followPointer: true
              },
              credits: {
                enabled: false,
              },
              plotOptions: {
                areaspline: {
                  fillOpacity: 0.1,
                },
                // series: {
                //   stickyTracking: true,
                //   events: {
                //     mouseOver: function () {
                //       var hoverSeries = this;

                //       this.chart.series.forEach(function (s) {
                //         if (s != hoverSeries) {
                //           // Turn off other series labels
                //           s.update({
                //             dataLabels: {
                //               enabled: false
                //             }
                //           });
                //         } else {
                //           // Turn on hover series labels
                //           hoverSeries.update({
                //             dataLabels: {
                //               enabled: true
                //             }
                //           });
                //         }
                //       });
                //     },
                //     mouseOut: function () {
                //       this.chart.series.forEach(function (s) {
                //         // Reset all series
                //         s.setState('');
                //         s.update({
                //           dataLabels: {
                //             enabled: true
                //           }
                //         });
                //       });
                //     }
                //   }
                // },
                enableMouseTracking: true,
                line: {
                  fillColor: '#FFFFFF',
                  dataLabels: {
                    enabled: false,
                    useHTML: true,

                    formatter: function () {

                      // if (this.y == 0) {
                      //   return this.x;
                      // }

                      // // if (this.y.toString().length < 3)
                      // //   return "<div> <div>" + this.y.toString() + ',0' + "</div>" + '<br>' + "<div>" + this.x + "</div>" + "</div>";
                      // // else
                      // //   return "<div> <div>" + this.y.toString().replace(".", ",") + "</div>" + '<br>' + "<div>" + this.x + "</div>" + "</div>";

                      // if (this.y.toString().length < 3)
                      //   return this.y.toString() + ',0';
                      // else
                      //   return this.y.toString().replace(".", ",");

                      // return "<div> <div>" + this.y.toString() + "</div>" + "<div>" + " <img class='logo-marcas' src='assets/imagePages/logo-pequeno.svg'>" + "</div>" + "</div>";
                      // return this.y.toString();

                      return "<div style='display:flex'>  <div>" + this.y.toString() + "</div>" + "<div style='displey:flex;' class='sig-negative'> <div>" + "<div>";

                    },
                    style: {
                      fontSize: '12px',
                      fontWeight: 'normal',
                      fontFamily: 'Poppins',
                      fontStyle: 'normal',
                      color: '#585656',
                    },
                  }
                },

              },
              series: this.graficoImagemPuraEvolutivoColuna1Model.Grafico,

              colors: cores,

              //   responsive: {
              //     rules: [{
              //         condition: {
              //             maxWidth: 500
              //         },
              //         chartOptions: {
              //             legend: {
              //                 layout: 'horizontal',
              //                 align: 'center',
              //                 verticalAlign: 'bottom'
              //             }
              //         }
              //     }]
              // },

              exporting: {
                enabled: false,
              },
              rules: [{
                condition: {
                  maxWidth: 1000
                },
                chartOptions: {
                  legend: {
                    layout: 'horizontal',
                    align: 'center',
                    verticalAlign: 'bottom',
                    horizontalAlign: 'center',
                  }
                }
              }]
            }
          }




          if (this.graficoImagemPuraEvolutivoColuna1Model?.Grafico?.length) {
            Highcharts.chart(graficoImagemPuraEvolutivoColuna1ModelHighchart).destroy();
            Highcharts.chart(graficoImagemPuraEvolutivoColuna1ModelHighchart);
          }


        }
      )


  }

  graficoLinha(nomeDiv, sequencia, marca: PadraoComboFiltro) {

    var graficoImagemPuraEvolutivoColunaModel = new GraficoLinhasModel();

    var textoTooltipPerc = this.translate.instant('grafico-texto-tooltip-perc');
    var textoTooltipMedia = this.translate.instant('grafico-texto-tooltip-media');
    var textoTooltipBase = this.translate.instant('grafico-texto-tooltip-base');

    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Marca.push(marca);
    filtros.Sequencia = sequencia;

    this.dashBoardNineService.ImagemEvolutiva(filtros)
      .subscribe((response: GraficoLinhasModel) => {

        graficoImagemPuraEvolutivoColunaModel = response;

        this.verificaBaseGraficoLinha(graficoImagemPuraEvolutivoColunaModel);

      }, (error) => console.error(error),
        () => {

          var grafico = [] as any;
          var index = 0;

          //var cores = this.getArrayColors(graficoImagemPuraEvolutivoColunaModel.Grafico.length);
          var cores = graficoImagemPuraEvolutivoColunaModel.Cores;
          // COR CINZA #75787c

          if (graficoImagemPuraEvolutivoColunaModel.Grafico?.length) {

            var lista = graficoImagemPuraEvolutivoColunaModel.Grafico.reverse();

            var objeto = null;
            var legenda = null;

            lista?.forEach(((item) => {

              if (item.tipoDados == 1) {
                objeto = {
                  id: item.name,
                  linkedTo: "legenda" + index,
                  type: 'line',
                  name: item.name,
                  data: item.data,
                  showInLegend: false,
                  marker: { "enabled": true, symbol: "circle", },
                  dataLabels: {
                    enabled: false,
                  },
                  // dashStyle: "ShortDash"
                }

                legenda = {
                  id: "legenda" + index,
                  name: item.name,
                  data: [],
                  color: cores[index],
                  type: 'line',
                  marker: { "enabled": true, symbol: "circle", },
                  // dashStyle: "ShortDash"

                }

              }

              else {

                objeto = {
                  id: item.name,
                  linkedTo: "legenda" + index,
                  type: 'line',
                  name: item.name,
                  data: item.data,
                  showInLegend: true,
                  marker: {
                    symbol: "circle",
                    // fillColor: '#FFFFFF', // BOLINHA VAZADA
                    lineWidth: 2,
                    radius: 4,
                    lineColor: cores[index]
                  },
                  dataLabels: {
                    // allowOverlap: true,
                    enabled: true
                  },
                }

                legenda = {
                  id: "legenda" + index,
                  name: item.name,
                  data: [],
                  color: cores[index],
                  type: 'column',
                  marker: { symbol: "circle", height: 10 }
                }

              }

              grafico.push(objeto);
              grafico.push(legenda);
              index++;
            }));

            graficoImagemPuraEvolutivoColunaModel.Grafico = grafico;

            var graficoImagemPuraEvolutivoColunaModelHighchart: any = {
              chart: {
                renderTo: nomeDiv,
                inverted: true,
                spacingTop: 0,
                height: 400,
                width: 87,
                title: {
                  text: "Texto",
                  enabled: false,
                },
                labels: {
                  enabled: true,
                },
                type: "spline",
                // //backgroundColor: "#F8F8F8",
                style: {
                  fontFamily: "Poppins",
                },
              },
              title: {
                text: "",
                enabled: false,
              },
              legend: {
                align: "left",
                verticalAlign: "top",
                enabled: false,
                alignColumns: true,
                itemDistance: 30,
                spacingBottom: 0,

                itemStyle: {
                  fontSize: '14px',
                  fontWeight: 'normal',
                  fontFamily: 'Poppins',
                  fontStyle: 'normal',
                  color: '#585656',
                  lineHeight: '20px',
                  textAlign: 'left',
                },

              },
              xAxis: [{
                lineWidth: 1,
                gridLineColor: '#ffffff',
                lineColor: '#ffffff',
                gridLineWidth: 1,
                tickInterval: 1,
                // opposite: true,
                // reversed: false,
                // reversedStacks: false,
                categories: graficoImagemPuraEvolutivoColunaModel.Periodos,
                labels: {
                  enabled: false,
                  floating: true,
                  // rotation: 90,
                  // align: 'left',
                  style: {
                    fontSize: '14px',
                    fontWeight: 'normal',
                    fontFamily: 'Poppins',
                    fontStyle: 'normal',
                    color: '#585656',
                    lineHeight: '20px',
                    textAlign: 'center',
                    width: 600,
                  },
                },


              }, {
                // top: '50%',
                // height: '50%',
                gridLineColor: '#ffffff',
                lineColor: '#ffffff',
                tickInterval: 1,
                reversed: false,
                reversedStacks: true,
                categories: graficoImagemPuraEvolutivoColunaModel.Periodos,
                labels: {
                  enabled: true,
                  floating: true,
                  rotation: 90,
                  align: 'left',
                  style: {
                    fontSize: '14px',
                    fontWeight: 'normal',
                    fontFamily: 'Poppins',
                    fontStyle: 'normal',
                    color: '#585656',
                    lineHeight: '20px',
                    textAlign: 'center',
                    width: 600,
                  },
                },
              }],
              yAxis: {
                lineWidth: 0,
                gridLineColor: '#ffffff',
                lineColor: '#ffffff',
                gridLineWidth: 0,
                tickInterval: 1,
                label: {
                  enabled: true,

                },
                title: {
                  text: "undef",
                  enabled: false,
                },
                labels: {
                  enabled: false,
                  align: 'right',
                  x: -2,

                  style: {
                    fontSize: '12px',
                    fontWeight: 'normal',
                    fontFamily: 'Poppins',
                    fontStyle: 'normal',
                    color: '#585656',
                  },
                },
              },

              tooltip: {
                shared: false,
                valueSuffix: "",
                enabled: true,
                useHTML: true,

                headerFormat: '',

                formatter(this: Highcharts.TooltipFormatterContextObject) {
                  var pointer = this.point as any;

                  console.log(pointer?.baseminima)

                  if (pointer?.baseminima != null && pointer?.baseminima != "") {
                    return '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> ' + textoTooltipMedia + '  </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.media + '</div>'
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px">  ' + textoTooltipPerc + ' </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.y + '%</div>  '
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> ' + textoTooltipBase + '</div>  <div style="font-weight: 400;text-align: right;color:red;">   ' + pointer?.valorbase + '</div> ';
                  }
                  else {
                    return '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> ' + textoTooltipMedia + '  </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.media + '</div>'
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px">  ' + textoTooltipPerc + ' </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.y + '%</div>  '
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> ' + textoTooltipBase + '</div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.valorbase + '</div> ';

                  }
                },

                // pointFormat:

                //   this.textoPoupUpGraficoLinha,

                outside: true,
                // backgroundColor: "rgba(246, 246, 246, 1)",
                // borderRadius: 30,
                // borderColor: "#bbbbbb",
                // borderWidth: 1.5,
                // style: { opacity: 1, background: "rgba(246, 246, 246, 1)" },
                // followPointer: true
              },
              credits: {
                enabled: false,
              },
              plotOptions: {
                areaspline: {
                  fillOpacity: 0.1,
                },
                series: {
                  stickyTracking: false,
                  events: {
                    mouseOver: function () {
                      var hoverSeries = this;

                      this.chart.series.forEach(function (s) {
                        if (s != hoverSeries) {
                          // Turn off other series labels
                          s.update({
                            dataLabels: {
                              enabled: false
                            }
                          });
                        } else {
                          // Turn on hover series labels
                          hoverSeries.update({
                            dataLabels: {
                              enabled: true
                            }
                          });
                        }
                      });
                    },
                    mouseOut: function () {
                      this.chart.series.forEach(function (s) {
                        // Reset all series
                        if (s.data.tipoDados == 1) {
                          s.setState('');
                          s.update({

                            dataLabels: {
                              enabled: true
                            }
                          });
                        }
                        // else{
                        //   s.setState('');
                        //   s.update({

                        //     dataLabels: {
                        //       enabled: false
                        //     }
                        //   });
                        // }
                      });


                    }
                  }
                },

                // series: [{
                //   data: ((arr, len) => {
                //     var i;
                //     for (i = 0; i < len; i = i + 1) {
                //       arr.push(i);
                //     }
                //     return arr;
                //   })([], 50),
                //   dataLabels: {
                //     allowOverlap: true,
                //     enabled: true,
                //     y: -5
                //   }
                // }],

                enableMouseTracking: true,
                line: {

                  dataLabels: {
                    enabled: true,
                    useHTML: true,
                    align: "center",
                    verticalAlign: "top",

                    alignColumns: true,
                    itemDistance: 40,
                    spacingBottom: 0,


                    height: 30,
                    width: 50,

                    formatter: function () {

                      // if (this.y == 0) {
                      //   return this.x;
                      // }

                      // // if (this.y.toString().length < 3)
                      // //   return "<div> <div>" + this.y.toString() + ',0' + "</div>" + '<br>' + "<div>" + this.x + "</div>" + "</div>";
                      // // else
                      // //   return "<div> <div>" + this.y.toString().replace(".", ",") + "</div>" + '<br>' + "<div>" + this.x + "</div>" + "</div>";

                      // if (this.y.toString().length < 3)
                      //   return this.y.toString() + ',0';
                      // else
                      //   return this.y.toString().replace(".", ",");

                      // return "<div> <div>" + this.y.toString() + "</div>" + "<div>" + " <img class='logo-marcas' src='assets/imagePages/logo-pequeno.svg'>" + "</div>" + "</div>";
                      // return this.y.toString();

                      var pointer = this.point as any;

                      // return this.y.toString();


                      if (pointer.sig != "") {
                        if (pointer.sig == "MAIOR") {
                          return "<div style='margin-top: -10px !important;display:flex;width: auto;justify-content: end;'>  <div style='display:flex;width: 35px'>" + this.y.toString() + "</div>" + "<div style='displey:flex;' class='sig-positive'> <div>" + "<div>";
                        }
                        else if (pointer.sig == "MENOR") {
                          return "<div style='margin-top: -10px !important;display:flex;width: auto;justify-content: end;'>  <div style='display:flex;width: 35px'>" + this.y.toString() + "</div>" + "<div style='displey:flex;' class='sig-negative'> <div>" + "<div>";
                        }
                        else
                          return "<div style='margin-top: -10px !important;display:flex;width: auto;justify-content: end;'>  <div style='display:flex;width: 35px'>" + this.y.toString() + "</div>" + "<div style='displey:flex;' class='sig-vazio'> <div>" + "<div>";
                      }
                      else
                        return "<div style='margin-top: -10px !important;display:flex;width: auto;justify-content: end;'>  <div style='display:flex;width: 35px'>" + this.y.toString() + "</div>" + "<div style='displey:flex;' class='sig-vazio'> <div>" + "<div>";

                    },
                    style: {
                      fontSize: '12px',
                      fontWeight: 'normal',
                      fontFamily: 'Poppins',
                      fontStyle: 'normal',
                      color: '#585656',

                    },
                  }
                },

              },
              series: graficoImagemPuraEvolutivoColunaModel.Grafico,

              colors: cores,



              //   responsive: {
              //     rules: [{
              //         condition: {
              //             maxWidth: 500
              //         },
              //         chartOptions: {
              //             legend: {
              //                 layout: 'horizontal',
              //                 align: 'center',
              //                 verticalAlign: 'bottom'
              //             }
              //         }
              //     }]
              // },

              exporting: {
                enabled: false,
              },
              rules: [{
                condition: {
                  maxWidth: 900
                },
                chartOptions: {
                  legend: {
                    layout: 'horizontal',
                    align: 'center',
                    verticalAlign: 'bottom',
                    horizontalAlign: 'center',
                  }
                }
              }]
            }
          }

          if (graficoImagemPuraEvolutivoColunaModel?.Grafico?.length) {
            Highcharts.chart(graficoImagemPuraEvolutivoColunaModelHighchart).destroy();
            Highcharts.chart(graficoImagemPuraEvolutivoColunaModelHighchart);
          }


        }
      )


  }

  verificaBaseEvolutivoMarcas(grafico: GraficoLinhasModel) {
    this.descVerificaBaseEvolutivoMarcas = "";
    if (grafico) {
      grafico.Grafico.forEach(item => {
        item.data.forEach(obj => {
          if (obj.baseminima != "") {
            this.descVerificaBaseEvolutivoMarcas = obj.baseminima;
          }
        });
      });
    }
  }

  verificaBaseEvolutivoMarcas2(grafico: GraficoLinhasModel) {
    this.descVerificaBaseEvolutivoMarcas2 = "";
    if (grafico) {
      grafico.Grafico.forEach(item => {
        item.data.forEach(obj => {
          if (obj.baseminima != "") {
            this.descVerificaBaseEvolutivoMarcas2 = obj.baseminima;
          }
        });
      });
    }
  }

  carregarEvolutivoMarcas() {


    var textoTooltipPerc = this.translate.instant('grafico-texto-tooltip-perc');
    var textoTooltipMedia = this.translate.instant('grafico-texto-tooltip-media');
    var textoTooltipBase = this.translate.instant('grafico-texto-tooltip-base');
    var textoTooltipPeriodo = this.translate.instant('grafico-texto-tooltip-periodo');

    var graficoEvolutivoMarcasLinhaModel = new GraficoLinhasModel();

    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Atributo = this.ModelAtributo.IdItem;

    if (!this.marcasSelecionadasGrafico3.length) {
      this.marcasSelecionadasGrafico3.push(this.filtroService.listaMarcas[0])
      this.marcasSelecionadasGrafico3.push(this.filtroService.listaMarcas[1])
      this.marcasSelecionadasGrafico3.push(this.filtroService.listaMarcas[2])
      this.marcasSelecionadasGrafico3.push(this.filtroService.listaMarcas[3])

    }

    // filtros.Marca = this.marcasSelecionadasGrafico3;
    filtros.Sequencia = 6;

    this.dashBoardNineService.ImagemEvolutivaLinhas(filtros)
      .subscribe((response: GraficoLinhasModel) => {

        this.ativaEvolutivoMarcasLinhas = true;
        graficoEvolutivoMarcasLinhaModel = response;

        this.verificaBaseEvolutivoMarcas(graficoEvolutivoMarcasLinhaModel);


      }, (error) => console.error(error),
        () => {

          var grafico = [] as any;
          var index = 0;
          var cores = this.getArrayColors(graficoEvolutivoMarcasLinhaModel.Grafico.length);
          //var cores = graficoEvolutivoMarcasLinhaModel.Cores;

          if (graficoEvolutivoMarcasLinhaModel?.Grafico?.length) {

            graficoEvolutivoMarcasLinhaModel.Grafico?.forEach(((item) => {

              var objeto = {
                id: item.name,
                linkedTo: "legenda" + index,
                type: 'line',
                name: item.name,
                data: item.data,
                showInLegend: false,
                marker: {
                  symbol: "circle",
                  // fillColor: '#FFFFFF',
                  lineWidth: 4,
                  radius: 10,
                  lineColor: cores[index]
                },

                dataLabels: {
                  allowOverlap: false,
                  enabled: true
                },

              }

              // var legenda = {
              //   id: "legenda" + index,
              //   name: item.name,
              //   data: [],
              //   color: cores[index],
              //   type: 'column',
              //   marker: {
              //     symbol: "circle", height: 10
              //   }
              //   // ,dataLabels: {
              //   //    allowOverlap: true,
              //   //   enabled: true
              //   // },
              // }

              var legenda = {
                id: "legenda" + index,
                name: item.name,
                data: [],
                color: cores[index],
                type: 'line',
                marker: {
                  symbol: "circle", height: 10,
                  width: 10,
                  lineWidth: 4,
                  radius: 50,
                },
                // dataLabels: {
                //   allowOverlap: true,
                //    enabled: true
                //  },
              }

              grafico.push(objeto);
              grafico.push(legenda);
              index++;
            }));

            graficoEvolutivoMarcasLinhaModel.Grafico = grafico;

            var graficoEvolutivoMarcasLinhaModelHighchart: any = {
              chart: {
                renderTo: 'container-evolutivo-marcas-linha',
                // inverted: true,
                spacingTop: 10,
                height: 450,
                // width: 100,
                title: {
                  text: "Texto",
                  enabled: false,
                },
                labels: {
                  enabled: true,
                },
                type: "spline",

                // //backgroundColor: "#F8F8F8",
                style: {
                  fontFamily: "Poppins",
                },
              },

              title: {
                text: "",
                enabled: false,
              },
              legend: {
                align: "left",
                verticalAlign: "bottom",
                enabled: true,
                alignColumns: true,
                itemDistance: 30,
                spacingBottom: 0,

                itemStyle: {
                  fontSize: '14px',
                  fontWeight: 'normal',
                  fontFamily: 'Poppins',
                  fontStyle: 'normal',
                  color: '#585656',
                  lineHeight: '20px',
                  textAlign: 'center',
                },

              },
              xAxis: [{
                lineWidth: 1,
                // gridLineColor: '#ffffff',
                // lineColor: '#ffffff',
                // gridLineWidth: 1,
                tickInterval: 1,
                // opposite: true,
                // reversed: false,
                // reversedStacks: false,




                categories: graficoEvolutivoMarcasLinhaModel.Periodos,

                labels: {
                  enabled: true,
                  floating: false,
                  // rotation: 90,
                  // align: 'left',
                  style: {
                    fontSize: '14px',
                    fontWeight: 'normal',
                    fontFamily: 'Poppins',
                    fontStyle: 'normal',
                    color: '#585656',
                    lineHeight: '20px',
                    textAlign: 'center',
                    width: 30,
                  },
                  formatter: function () {

                    // var valorBase = graficoEvolutivoMarcasLinhaModel.Bases[this.pos];

                    return '<div style="font-style: normal;font-weight: 500;font-size: 12px;line-height: 20px;text-align: center;color: #1B212D;">' + this.value + '</div> '
                    // + '<div style="font-style: normal;font-weight: 400;font-size: 12px;line-height: 20px;text-align: center;color: #6E778B;">' + valorBase + '</div>';
                    // return this.value + " " + item;
                  },

                },


              },
              ],
              yAxis: {
                lineWidth: 0,
                gridLineColor: '#ffffff',
                lineColor: '#ffffff',
                gridLineWidth: 0,
                tickInterval: 2,

                label: {
                  enabled: false,

                },
                title: {
                  text: "undef",
                  enabled: false,
                },
                labels: {

                  enabled: true,
                  align: 'center',
                  // x: 22,

                  style: {
                    fontSize: '12px',
                    fontWeight: 'normal',
                    fontFamily: 'Poppins',
                    fontStyle: 'normal',
                    color: '#585656',
                  },
                },
              },

              tooltip: {
                shared: false,
                valueSuffix: "",
                enabled: true,
                useHTML: true,

                headerFormat: '',
                width: 250,
                formatter(this: Highcharts.TooltipFormatterContextObject) {
                  var pointer = this.point as any;
                  if (pointer?.baseminima != null && pointer?.baseminima != "") {
                    return '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> ' + textoTooltipPerc + '  </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.y.toFixed(1) + '%</div>'
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px">  ' + textoTooltipBase + ' </div>  <div style="font-weight: 400;text-align: right;color:red;">   ' + pointer?.valorbase + '</div>  '
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> ' + textoTooltipPeriodo + '</div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.periodo + '</div> ';
                  }
                  else {
                    return '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> ' + textoTooltipPerc + '  </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.y.toFixed(1) + '%</div>'
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px">  ' + textoTooltipBase + ' </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.valorbase + '</div>  '
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> ' + textoTooltipPeriodo + '</div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.periodo + '</div> ';
                  }
                  // if (pointer?.valorbase == 0) {
                  //   return '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: center;"> Média ' + pointer?.y.toFixed(1) + '% </div>  <br/>'
                  //     + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: center;"> ' + pointer?.valorbase + '</div>  <br/> '
                  //     + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: center;"> Período ' + pointer?.periodo + '</div>  <br/>';
                  // } else {
                  //   return '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 118px"> Média  </div>  <div style="font-weight: 400;text-align: center;">' + pointer?.y.toFixed(1) + '%</div>  <br/>'
                  //     + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 118px">  Base </div>  <div style="font-weight: 400;text-align: center;">' + pointer?.valorbase + '</div>  <br/> '
                  //     + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 118px"> Período </div>  <div style="font-weight: 400;text-align: center;">' + pointer?.periodo + '</div>  <br/>';
                  // };
                },

                // pointFormat:

                //   this.textoPoupUpGraficoLinha,

                // outside: true,
                // backgroundColor: "rgba(246, 246, 246, 1)",
                // borderRadius: 30,
                // borderColor: "#bbbbbb",
                // borderWidth: 1.5,
                // style: { opacity: 1, background: "rgba(246, 246, 246, 1)" },
                // followPointer: true
              },
              credits: {
                enabled: false,
              },
              plotOptions: {
                areaspline: {
                  fillOpacity: 0.1,
                },
                series: {
                  //  lineWidth: 1,
                  allowPointSelect: true,
                  dataLabels: {
                    enabled: true,
                    inside: true,
                  },
                  stickyTracking: false,
                  events: {
                    mouseOver: function () {
                      var hoverSeries = this;
                      // Evita que o grafico trave por ecesso de chamadas so evento
                      if (graficoEvolutivoMarcasLinhaModel.Grafico.length < 60)
                        this.chart.series.forEach(function (s) {
                          if (s != hoverSeries) {
                            // Turn off other series labels
                            s.update({
                              dataLabels: {
                                enabled: false
                              }
                            });
                          } else {
                            // Turn on hover series labels
                            hoverSeries.update({
                              dataLabels: {
                                enabled: true
                              }
                            });
                          }
                        });
                    },

                    mouseOut: function () {
                      // Evita que o grafico trave por ecesso de chamadas so evento
                      if (graficoEvolutivoMarcasLinhaModel.Grafico.length < 60)
                        this.chart.series.forEach(function (s) {
                          // Reset all series
                          s.setState('');
                          s.update({
                            dataLabels: {
                              enabled: true
                            }
                          });
                        });
                    }
                  }
                },

                enableMouseTracking: true,
                line: {
                  enabled: true,
                  allowOverlap: false,

                  dataLabels: {
                    enabled: false,
                    useHTML: true,
                    y: 14,
                    x: -1,
                    formatter: function () {

                      var pointer = this.point as any;
                      if (pointer.sig != "") {
                        if (pointer.sig == "MAIOR") {
                          return "<div style='display:flex;'>  <div>" + this.y.toString() + "</div>" + "<div style='display:flex;position: absolute;left: 18px;bottom: 4px;' class='sig-positive'> <div>" + "<div>";
                        }
                        else if (pointer.sig == "MENOR") {
                          return "<div style='display:flex;'>  <div>" + this.y.toString() + "</div>" + "<div style='display:flex;position: absolute;left: 18px;bottom: 4px;' class='sig-negative'> <div>" + "<div>";
                        }
                        else
                          return "<div style='display:flex;'>  <div>" + this.y.toString() + "</div>" + "<div style='display:flex;position: absolute;left: 18px;bottom: 4px;' class='sig-vazio'> <div>" + "<div>";
                      }
                      else
                        return "<div style='display:flex;'>  <div>" + this.y.toString() + "</div>" + "<div style='display:flex;position: absolute;left: 18px;bottom: 4px;' class='sig-vazio'> <div>" + "<div>";

                    },
                    align: 'center',

                    style: {
                      fontSize: '11px',
                      fontWeight: 'normal',
                      fontFamily: 'Poppins',
                      fontStyle: 'normal',
                      color: '#FEFEFE',
                      textAlign: 'center',
                      // width: 30,
                    },
                  }
                },

              },



              // marker: {
              //   radius: 15,
              //   fillColor: '#FFF',
              //   lineColor: '#7cb5ec',
              //   lineWidth: 4
              // },
              series: graficoEvolutivoMarcasLinhaModel.Grafico,

              colors: cores,



              //   responsive: {
              //     rules: [{
              //         condition: {
              //             maxWidth: 500
              //         },
              //         chartOptions: {
              //             legend: {
              //                 layout: 'horizontal',
              //                 align: 'center',
              //                 verticalAlign: 'bottom'
              //             }
              //         }
              //     }]
              // },

              exporting: {
                enabled: false,
              },
              rules: [{
                condition: {
                  maxWidth: 900
                },
                chartOptions: {
                  // legend: {
                  //   layout: 'horizontal',
                  //   align: 'center',
                  //   verticalAlign: 'bottom',
                  //   horizontalAlign: 'center',
                  // }
                }
              }]
            }
          }
          else{
            this.ativaEvolutivoMarcasLinhas = true;
          }


          if (graficoEvolutivoMarcasLinhaModel?.Grafico?.length) {
            Highcharts.chart(graficoEvolutivoMarcasLinhaModelHighchart).destroy();
            Highcharts.chart(graficoEvolutivoMarcasLinhaModelHighchart);
          }

        }
      )
  }



  descInicial: string = "Nestlé";
  urlImage: string;
  public ModelMarcas: PadraoComboFiltro;
  listaMarcas: Array<PadraoComboFiltro>;
  montaImageMarca() {

    if (this.ModelMarcas != undefined) {
      this.descInicial = this.ModelMarcas.DescItem;
      this.urlImage = "assets/marcas/" + this.ModelMarcas.IdItem + ".svg";
    }
    else {
      this.descInicial = this.listaMarcas[0].DescItem;
      this.urlImage = "assets/marcas/" + this.listaMarcas[0].IdItem + ".svg";
    }
  }

  onchangeMarcaGrafico(item: PadraoComboFiltro) {
    this.ModelMarcas = item;
    this.montaImageMarca();
    this.carregarEvolutivoMarcas2();
  }
  carregarEvolutivoMarcas2() {

    var textoTooltipPerc = this.translate.instant('grafico-texto-tooltip-perc');
    var textoTooltipMedia = this.translate.instant('grafico-texto-tooltip-media');
    var textoTooltipBase = this.translate.instant('grafico-texto-tooltip-base');
    var textoTooltipPeriodo = this.translate.instant('grafico-texto-tooltip-periodo');

    var graficoEvolutivoMarcasLinhaModel = new GraficoLinhasModel();

    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Atributo = 11;
    var list = [];
    list.push(this.ModelMarcas);
    filtros.Marca = list;

    filtros.Sequencia = 14;

    this.dashBoardNineService.ImagemEvolutivaLinhas2(filtros)
      .subscribe((response: GraficoLinhasModel) => {

        this.ativaEvolutivoMarcasLinhas2 = true;
        graficoEvolutivoMarcasLinhaModel = response;

        this.verificaBaseEvolutivoMarcas2(graficoEvolutivoMarcasLinhaModel);

      }, (error) => console.error(error),
        () => {

          var grafico = [] as any;
          var index = 0;
          // var cores = graficoEvolutivoMarcasLinhaModel.Cores;
          var cores = this.getArrayColors(graficoEvolutivoMarcasLinhaModel.Grafico.length);

          if (graficoEvolutivoMarcasLinhaModel?.Grafico?.length) {

            graficoEvolutivoMarcasLinhaModel.Grafico?.forEach(((item) => {

              var objeto = {
                id: item.name,
                linkedTo: "legenda" + index,
                type: 'line',
                name: item.name,
                data: item.data,
                showInLegend: false,
                marker: {
                  symbol: "circle",
                  // fillColor: '#FFFFFF',
                  lineWidth: 4,
                  radius: 10,
                  lineColor: cores[index]
                },

                dataLabels: {
                  allowOverlap: false,
                  enabled: true
                },
              }

              var legenda = {
                id: "legenda" + index,
                name: item.name,
                data: [],
                color: cores[index],
                type: 'line',
                marker: {
                  symbol: "circle", height: 10,
                  width: 10,
                  lineWidth: 4,
                  radius: 50,
                },
              }

              grafico.push(objeto);
              grafico.push(legenda);
              index++;
            }));

            graficoEvolutivoMarcasLinhaModel.Grafico = grafico;

            var graficoEvolutivoMarcasLinhaModelHighchart: any = {
              chart: {
                renderTo: 'container-evolutivo-marcas-linha-2',
                // inverted: true,
                spacingTop: 10,
                height: 450,
                // width: 100,
                title: {
                  text: "Texto",
                  enabled: false,
                },
                labels: {
                  enabled: true,
                },
                type: "spline",

                // //backgroundColor: "#F8F8F8",
                style: {
                  fontFamily: "Poppins",
                },
              },

              title: {
                text: "",
                enabled: false,
              },
              legend: {
                align: "left",
                verticalAlign: "bottom",
                enabled: true,
                alignColumns: true,
                itemDistance: 30,
                spacingBottom: 0,

                itemStyle: {
                  fontSize: '14px',
                  fontWeight: 'normal',
                  fontFamily: 'Poppins',
                  fontStyle: 'normal',
                  color: '#585656',
                  lineHeight: '20px',
                  textAlign: 'center',
                },

              },
              xAxis: [{
                lineWidth: 1,
                // gridLineColor: '#ffffff',
                // lineColor: '#ffffff',
                // gridLineWidth: 1,
                tickInterval: 1,
                // opposite: true,
                // reversed: false,
                // reversedStacks: false,
                categories: graficoEvolutivoMarcasLinhaModel.Periodos,

                labels: {
                  enabled: true,
                  floating: false,
                  // rotation: 90,
                  // align: 'left',
                  style: {
                    fontSize: '14px',
                    fontWeight: 'normal',
                    fontFamily: 'Poppins',
                    fontStyle: 'normal',
                    color: '#585656',
                    lineHeight: '20px',
                    textAlign: 'center',
                    width: 30,
                  },
                  formatter: function () {
                    return '<div style="font-style: normal;font-weight: 500;font-size: 12px;line-height: 20px;text-align: center;color: #1B212D;">' + this.value + '</div> '
                  },

                },


              },
              ],
              yAxis: {
                lineWidth: 0,
                gridLineColor: '#ffffff',
                lineColor: '#ffffff',
                gridLineWidth: 0,
                tickInterval: 2,

                label: {
                  enabled: false,

                },
                title: {
                  text: "undef",
                  enabled: false,
                },
                labels: {

                  enabled: true,
                  align: 'center',
                  // x: 22,

                  style: {
                    fontSize: '12px',
                    fontWeight: 'normal',
                    fontFamily: 'Poppins',
                    fontStyle: 'normal',
                    color: '#585656',
                  },
                },
              },

              tooltip: {
                shared: false,
                valueSuffix: "",
                enabled: true,
                useHTML: true,

                headerFormat: '',
                width: 250,
                formatter(this: Highcharts.TooltipFormatterContextObject) {
                  var pointer = this.point as any;
                  if (pointer?.baseminima != null && pointer?.baseminima != "") {
                    return '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> ' + textoTooltipPerc + '  </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.y.toFixed(1) + '%</div>'
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px">  ' + textoTooltipBase + ' </div>  <div style="font-weight: 400;text-align: right;color:red;">   ' + pointer?.valorbase + '</div>  '
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> ' + textoTooltipPeriodo + '</div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.periodo + '</div> ';
                  }
                  else {
                    return '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> ' + textoTooltipPerc + '  </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.y.toFixed(1) + '%</div>'
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px">  ' + textoTooltipBase + ' </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.valorbase + '</div>  '
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> ' + textoTooltipPeriodo + '</div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.periodo + '</div> ';
                  }

                },

              },
              credits: {
                enabled: false,
              },
              plotOptions: {
                areaspline: {
                  fillOpacity: 0.1,
                },
                series: {
                  //  lineWidth: 1,
                  allowPointSelect: true,
                  dataLabels: {
                    enabled: true,
                    inside: true,
                  },
                  stickyTracking: false,
                  events: {
                    mouseOver: function () {
                      var hoverSeries = this;
                      // Evita que o grafico trave por ecesso de chamadas so evento
                      if (graficoEvolutivoMarcasLinhaModel.Grafico.length < 60)
                        this.chart.series.forEach(function (s) {
                          if (s != hoverSeries) {
                            // Turn off other series labels
                            s.update({
                              dataLabels: {
                                enabled: false
                              }
                            });
                          } else {
                            // Turn on hover series labels
                            hoverSeries.update({
                              dataLabels: {
                                enabled: true
                              }
                            });
                          }
                        });
                    },

                    mouseOut: function () {
                      // Evita que o grafico trave por ecesso de chamadas so evento
                      if (graficoEvolutivoMarcasLinhaModel.Grafico.length < 60)
                        this.chart.series.forEach(function (s) {
                          // Reset all series
                          s.setState('');
                          s.update({
                            dataLabels: {
                              enabled: true
                            }
                          });
                        });
                    }
                  }
                },

                enableMouseTracking: true,
                line: {
                  enabled: true,
                  allowOverlap: false,

                  dataLabels: {
                    enabled: false,
                    useHTML: true,
                    y: 14,
                    x: -1,
                    formatter: function () {

                      var pointer = this.point as any;
                      if (pointer.sig != "") {
                        if (pointer.sig == "MAIOR") {
                          return "<div style='display:flex;'>  <div>" + this.y.toString() + "</div>" + "<div style='display:flex;position: absolute;left: 18px;bottom: 4px;' class='sig-positive'> <div>" + "<div>";
                        }
                        else if (pointer.sig == "MENOR") {
                          return "<div style='display:flex;'>  <div>" + this.y.toString() + "</div>" + "<div style='display:flex;position: absolute;left: 18px;bottom: 4px;' class='sig-negative'> <div>" + "<div>";
                        }
                        else
                          return "<div style='display:flex;'>  <div>" + this.y.toString() + "</div>" + "<div style='display:flex;position: absolute;left: 18px;bottom: 4px;' class='sig-vazio'> <div>" + "<div>";
                      }
                      else
                        return "<div style='display:flex;'>  <div>" + this.y.toString() + "</div>" + "<div style='display:flex;position: absolute;left: 18px;bottom: 4px;' class='sig-vazio'> <div>" + "<div>";

                    },
                    align: 'center',

                    style: {
                      fontSize: '11px',
                      fontWeight: 'normal',
                      fontFamily: 'Poppins',
                      fontStyle: 'normal',
                      color: '#FEFEFE',
                      textAlign: 'center',
                      // width: 30,
                    },
                  }
                },

              },
              series: graficoEvolutivoMarcasLinhaModel.Grafico,
              colors: cores,
              exporting: {
                enabled: false,
              },
              rules: [{
                condition: {
                  maxWidth: 900
                },
                chartOptions: {
                  // legend: {
                  //   layout: 'horizontal',
                  //   align: 'center',
                  //   verticalAlign: 'bottom',
                  //   horizontalAlign: 'center',
                  // }
                }
              }]
            }
          }

          if (graficoEvolutivoMarcasLinhaModel?.Grafico?.length) {
            Highcharts.chart(graficoEvolutivoMarcasLinhaModelHighchart).destroy();
            Highcharts.chart(graficoEvolutivoMarcasLinhaModelHighchart);
          }

        }
      )
  }


  carregarTabelaAdHocAtributo() {

    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Sequencia = 7;
    filtros.Marca.push(this.marcatabelaAdHoc)

    this.dashBoardNineService.ImagemEvolutiTabelaAdHocAtributovaLinhas2(filtros)
      .subscribe((response: TabelaPadraoAdHoc) => {
        if (response) {
          this.ativaTabelaAdHoc = true;
          var titulos = response.Titulos;
          var dados = response.Dados;
          this.tabelaPadraoAdHoc.Dados = dados;
          this.tabelaPadraoAdHoc.Titulos = titulos;
        }

      }, (error) => console.error(error),
        () => {

        })
  }

  TabelaAdHocAtributoBloco6() {

    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Sequencia = 7;
    filtros.Marca.push(this.marcatabelaAdHoc)

    this.dashBoardNineService.TabelaAdHocAtributoBloco6(filtros)
      .subscribe((response: TabelaPadraoAdHoc) => {
        if (response) {
          this.ativaTabelaAdHocBloco6 = true;
          var titulos = response.Titulos;
          var dados = response.Dados;
          this.tabelaPadraoAdHocBloco6.Dados = dados;
          this.tabelaPadraoAdHocBloco6.Titulos = titulos;
        }

      }, (error) => console.error(error),
        () => {

        })
  }


  TabelaAdHocAtributoBloco10() {

    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Sequencia = 8;
    filtros.Marca.push(this.marcatabelaAdHocBloco10)

    this.dashBoardNineService.TabelaAdHocAtributoBloco10(filtros)
      .subscribe((response: TabelaPadraoAdHoc) => {
        if (response) {
          this.ativaTabelaAdHocBloco10 = true;
          debugger
          var titulos = response.Titulos;
          var dados = response.Dados;
          this.tabelaPadraoAdHocBloco10.Dados = dados;
          this.tabelaPadraoAdHocBloco10.Titulos = titulos;
        }

      }, (error) => console.error(error),
        () => {

        })
  }


  TabelaAdHocAtributoBloco2() {

    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Sequencia = 15;

    this.dashBoardNineService.TabelaAdHocAtributoBloco2(filtros)
      .subscribe((response: TabelaPadraoAdHoc) => {
        if (response) {
          this.ativaTabelaAdHocBloco2 = true;
          var titulos = response.Titulos;
          var dados = response.Dados;
          this.tabelaPadraoAdHocBloco2.Dados = dados;
          this.tabelaPadraoAdHocBloco2.Titulos = titulos;
        }

      }, (error) => console.error(error),
        () => {

        })
  }


  descInicial2: string = "Nestlé";
  urlImage2: string;
  public ModelMarcas2: PadraoComboFiltro;
  listaMarcas2: Array<PadraoComboFiltro>;
  montaImageMarca2() {

    if (this.ModelMarcas2 != undefined) {
      this.descInicial2 = this.ModelMarcas2.DescItem;
      this.urlImage2 = "assets/marcas/" + this.ModelMarcas2.IdItem + ".svg";
    }
    else {
      this.descInicial2 = this.listaMarcas2[0].DescItem;
      this.urlImage2 = "assets/marcas/" + this.listaMarcas2[0].IdItem + ".svg";
    }
  }

  onchangeMarcaGrafico2(item: PadraoComboFiltro) {
    this.ModelMarcas = item;
    this.montaImageMarca2();
    this.carregarEvolutivoMarcas02();
  }
  carregarEvolutivoMarcas02() {

    var textoTooltipPerc = this.translate.instant('grafico-texto-tooltip-perc');
    var textoTooltipMedia = this.translate.instant('grafico-texto-tooltip-media');
    var textoTooltipBase = this.translate.instant('grafico-texto-tooltip-base');
    var textoTooltipPeriodo = this.translate.instant('grafico-texto-tooltip-periodo');

    var graficoEvolutivoMarcasLinhaModel = new GraficoLinhasModel();

    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Atributo = 12;
    var list = [];
    list.push(this.ModelMarcas2);
    filtros.Marca = list;

    filtros.Sequencia = 16;

    this.dashBoardNineService.ImagemEvolutivaLinhas2(filtros)
      .subscribe((response: GraficoLinhasModel) => {

        this.ativaEvolutivoMarcasLinhas2 = true;
        graficoEvolutivoMarcasLinhaModel = response;

        this.verificaBaseEvolutivoMarcas2(graficoEvolutivoMarcasLinhaModel);

      }, (error) => console.error(error),
        () => {

          var grafico = [] as any;
          var index = 0;
          // var cores = graficoEvolutivoMarcasLinhaModel.Cores;
          var cores = this.getArrayColors(graficoEvolutivoMarcasLinhaModel.Grafico.length);

          if (graficoEvolutivoMarcasLinhaModel?.Grafico?.length) {

            graficoEvolutivoMarcasLinhaModel.Grafico?.forEach(((item) => {

              var objeto = {
                id: item.name,
                linkedTo: "legenda" + index,
                type: 'line',
                name: item.name,
                data: item.data,
                showInLegend: false,
                marker: {
                  symbol: "circle",
                  // fillColor: '#FFFFFF',
                  lineWidth: 4,
                  radius: 10,
                  lineColor: cores[index]
                },

                dataLabels: {
                  allowOverlap: false,
                  enabled: true
                },
              }

              var legenda = {
                id: "legenda" + index,
                name: item.name,
                data: [],
                color: cores[index],
                type: 'line',
                marker: {
                  symbol: "circle", height: 10,
                  width: 10,
                  lineWidth: 4,
                  radius: 50,
                },
              }

              grafico.push(objeto);
              grafico.push(legenda);
              index++;
            }));

            graficoEvolutivoMarcasLinhaModel.Grafico = grafico;

            var graficoEvolutivoMarcasLinhaModelHighchart: any = {
              chart: {
                renderTo: 'container-evolutivo-marcas-linha-22',
                // inverted: true,
                spacingTop: 10,
                height: 550,
                // width: 100,
                title: {
                  text: "Texto",
                  enabled: false,
                },
                labels: {
                  enabled: true,
                },
                type: "spline",

                // //backgroundColor: "#F8F8F8",
                style: {
                  fontFamily: "Poppins",
                },
              },

              title: {
                text: "",
                enabled: false,
              },
              legend: {
                align: "left",
                verticalAlign: "bottom",
                enabled: true,
                alignColumns: true,
                itemDistance: 30,
                spacingBottom: 0,

                itemStyle: {
                  fontSize: '14px',
                  fontWeight: 'normal',
                  fontFamily: 'Poppins',
                  fontStyle: 'normal',
                  color: '#585656',
                  lineHeight: '20px',
                  textAlign: 'center',
                },

              },
              xAxis: [{
                lineWidth: 1,
                // gridLineColor: '#ffffff',
                // lineColor: '#ffffff',
                // gridLineWidth: 1,
                tickInterval: 1,
                // opposite: true,
                // reversed: false,
                // reversedStacks: false,
                categories: graficoEvolutivoMarcasLinhaModel.Periodos,

                labels: {
                  enabled: true,
                  floating: false,
                  // rotation: 90,
                  // align: 'left',
                  style: {
                    fontSize: '14px',
                    fontWeight: 'normal',
                    fontFamily: 'Poppins',
                    fontStyle: 'normal',
                    color: '#585656',
                    lineHeight: '20px',
                    textAlign: 'center',
                    width: 30,
                  },
                  formatter: function () {
                    return '<div style="font-style: normal;font-weight: 500;font-size: 12px;line-height: 20px;text-align: center;color: #1B212D;">' + this.value + '</div> '
                  },

                },


              },
              ],
              yAxis: {
                lineWidth: 0,
                gridLineColor: '#ffffff',
                lineColor: '#ffffff',
                gridLineWidth: 0,
                tickInterval: 2,

                label: {
                  enabled: false,

                },
                title: {
                  text: "undef",
                  enabled: false,
                },
                labels: {

                  enabled: true,
                  align: 'center',
                  // x: 22,

                  style: {
                    fontSize: '12px',
                    fontWeight: 'normal',
                    fontFamily: 'Poppins',
                    fontStyle: 'normal',
                    color: '#585656',
                  },
                },
              },

              tooltip: {
                shared: false,
                valueSuffix: "",
                enabled: true,
                useHTML: true,

                headerFormat: '',
                width: 250,
                formatter(this: Highcharts.TooltipFormatterContextObject) {
                  var pointer = this.point as any;
                  if (pointer?.baseminima != null && pointer?.baseminima != "") {
                    return '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> ' + textoTooltipPerc + '  </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.y.toFixed(1) + '%</div>'
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px">  ' + textoTooltipBase + ' </div>  <div style="font-weight: 400;text-align: right;color:red;">   ' + pointer?.valorbase + '</div>  '
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> ' + textoTooltipPeriodo + '</div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.periodo + '</div> ';
                  }
                  else {
                    return '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> ' + textoTooltipPerc + '  </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.y.toFixed(1) + '%</div>'
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px">  ' + textoTooltipBase + ' </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.valorbase + '</div>  '
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px"> ' + textoTooltipPeriodo + '</div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.periodo + '</div> ';
                  }

                },

              },
              credits: {
                enabled: false,
              },
              plotOptions: {
                areaspline: {
                  fillOpacity: 0.1,
                },
                series: {
                  //  lineWidth: 1,
                  allowPointSelect: true,
                  dataLabels: {
                    enabled: true,
                    inside: true,
                  },
                  stickyTracking: false,
                  events: {
                    mouseOver: function () {
                      var hoverSeries = this;
                      // Evita que o grafico trave por ecesso de chamadas so evento
                      if (graficoEvolutivoMarcasLinhaModel.Grafico.length < 60)
                        this.chart.series.forEach(function (s) {
                          if (s != hoverSeries) {
                            // Turn off other series labels
                            s.update({
                              dataLabels: {
                                enabled: false
                              }
                            });
                          } else {
                            // Turn on hover series labels
                            hoverSeries.update({
                              dataLabels: {
                                enabled: true
                              }
                            });
                          }
                        });
                    },

                    mouseOut: function () {
                      // Evita que o grafico trave por ecesso de chamadas so evento
                      if (graficoEvolutivoMarcasLinhaModel.Grafico.length < 60)
                        this.chart.series.forEach(function (s) {
                          // Reset all series
                          s.setState('');
                          s.update({
                            dataLabels: {
                              enabled: true
                            }
                          });
                        });
                    }
                  }
                },

                enableMouseTracking: true,
                line: {
                  enabled: true,
                  allowOverlap: false,

                  dataLabels: {
                    enabled: false,
                    useHTML: true,
                    y: 14,
                    x: -1,
                    formatter: function () {

                      var pointer = this.point as any;
                      if (pointer.sig != "") {
                        if (pointer.sig == "MAIOR") {
                          return "<div style='display:flex;'>  <div>" + this.y.toString() + "</div>" + "<div style='display:flex;position: absolute;left: 18px;bottom: 4px;' class='sig-positive'> <div>" + "<div>";
                        }
                        else if (pointer.sig == "MENOR") {
                          return "<div style='display:flex;'>  <div>" + this.y.toString() + "</div>" + "<div style='display:flex;position: absolute;left: 18px;bottom: 4px;' class='sig-negative'> <div>" + "<div>";
                        }
                        else
                          return "<div style='display:flex;'>  <div>" + this.y.toString() + "</div>" + "<div style='display:flex;position: absolute;left: 18px;bottom: 4px;' class='sig-vazio'> <div>" + "<div>";
                      }
                      else
                        return "<div style='display:flex;'>  <div>" + this.y.toString() + "</div>" + "<div style='display:flex;position: absolute;left: 18px;bottom: 4px;' class='sig-vazio'> <div>" + "<div>";

                    },
                    align: 'center',

                    style: {
                      fontSize: '11px',
                      fontWeight: 'normal',
                      fontFamily: 'Poppins',
                      fontStyle: 'normal',
                      color: '#FEFEFE',
                      textAlign: 'center',
                      // width: 30,
                    },
                  }
                },

              },
              series: graficoEvolutivoMarcasLinhaModel.Grafico,
              colors: cores,
              exporting: {
                enabled: false,
              },
              rules: [{
                condition: {
                  maxWidth: 900
                },
                chartOptions: {
                  // legend: {
                  //   layout: 'horizontal',
                  //   align: 'center',
                  //   verticalAlign: 'bottom',
                  //   horizontalAlign: 'center',
                  // }
                }
              }]
            }
          }

          if (graficoEvolutivoMarcasLinhaModel?.Grafico?.length) {
            Highcharts.chart(graficoEvolutivoMarcasLinhaModelHighchart).destroy();
            Highcharts.chart(graficoEvolutivoMarcasLinhaModelHighchart);
          }

        }
      )
  }


  onchangeMarcaTabelaAdHocBloco6(item: PadraoComboFiltro) {
    this.marcatabelaAdHocBloco6 = item;
    this.TabelaAdHocAtributoBloco6();
  }

  onchangeMarcaTabelaAdHocBloco10(item: PadraoComboFiltro) {
    this.marcatabelaAdHocBloco10 = item;
    this.TabelaAdHocAtributoBloco10();
  }



  onchangeMarcaTabelaAdHoc(item: PadraoComboFiltro) {
    this.marcatabelaAdHoc = item;
    this.carregarTabelaAdHocAtributo();
  }


  onchangeMarcaColuna(item: PadraoComboFiltro, nCol: number) {
    switch (nCol) {
      case 1:
        this.marcaColuna1 = item;
        break;
      case 2:
        this.marcaColuna2 = item;
        break;
      case 3:
        this.marcaColuna3 = item;
        break;
      case 4:
        this.marcaColuna4 = item;
        break;
      case 5:
        this.marcaColuna5 = item;
        break;
    }

    this.graficoLinha("container-coluna-" + nCol.toString(), nCol, item);

  }

  ajustaScrollFinal() {
    if (document.getElementById('aqui_id_div_grafico') != null && !document.getElementById('aqui_id_div_grafico').scrollLeft) {
      document.getElementById('aqui_id_div_grafico').scrollLeft = 9999;
    }

  }



  async exportToExcelGrafico1() {
    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Marca = this.marcasSelecionadasGrafico3;
    filtros.Sequencia = 1;
    filtros.Atributo = this.ModelAtributo.IdItem;

    var titulo = this.translate.instant('dashboard-nine-grafico11-titulo');
    if (this.ModelAtributo) {
      titulo = this.ModelAtributo.DescItem;
    }

    filtros.TituloGrafico = titulo;

    await this.downloadArquivoService.
      DownloadDashboardNineComparativoImagemPura(filtros)
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
    var titulo = this.translate.instant('dashboard-nine-grafico12-titulo');
    filtroPadrao.TituloGrafico = titulo.substring(0, 10);
    await this.downloadArquivoService.
      DownloadDashboardNineEvolutivoPeriodos(filtroPadrao)
      .subscribe((result) => {

        saveAs(
          result,
          titulo.toString().substring(0, 15) + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => console.error(error));
  }

  async exportToExcelGrafico3() {
    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    var list = [];
    list.push(this.ModelMarcas);
    filtros.Marca = list;

    filtros.Sequencia = 1;
    filtros.Atributo = 11;

    var titulo = this.translate.instant('dashboard-nine-grafico14-titulo');
    filtros.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardNineComparativoImagemPura2(filtros)
      .subscribe((result) => {

        saveAs(
          result,
          titulo + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => console.error(error));
  }

  async exportToExcelGrafico03() {
    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    var list = [];
    list.push(this.ModelMarcas2);
    filtros.Marca = list;
    filtros.Sequencia = 1;
    filtros.Atributo = 12;

    var titulo = this.translate.instant('dashboard-nine-grafico16-titulo');
    filtros.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardNineComparativoImagemPura2(filtros)
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
    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Sequencia = 8;
    filtros.Marca.push(this.marcatabelaAdHoc)

    var titulo = this.translate.instant('dashboard-nine-grafico14-titulo');
    filtros.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardNineTabelaAdHoc(filtros)
      .subscribe((result) => {

        saveAs(
          result,
          titulo + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => console.error(error));
  }


  async exportToExcelGraficobloco6() {
    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Sequencia = 8;
    filtros.Marca.push(this.marcatabelaAdHoc)

    var titulo = this.translate.instant('dashboard-nine-grafico12-titulo');
    filtros.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardNineTabelaAdHocBloco6(filtros)
      .subscribe((result) => {

        saveAs(
          result,
          titulo + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => console.error(error));
  }


  async exportToExcelGraficobloco10() {
    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Sequencia = 8;
    filtros.Marca.push(this.marcatabelaAdHocBloco10)

    var titulo = this.translate.instant('dashboard-nine-grafico13-titulo');
    filtros.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardNineTabelaAdHocBloco10(filtros)
      .subscribe((result) => {

        saveAs(
          result,
          titulo + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => console.error(error));
  }


  async exportToExcelGraficobloco2() {
    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Sequencia = 15;

    var titulo = this.translate.instant('dashboard-nine-grafico15-titulo');
    filtros.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardNineTabelaAdHocBloco2(filtros)
      .subscribe((result) => {

        saveAs(
          result,
          titulo.substring(0, 35) + "..." + '_' +
          this.downloadArquivoService.getData() +
          ".xlsx"
        );
      }, (error) => console.error(error));
  }




  exportToPptComparativoExperiencia() {
    var nome = this.translate.instant('dashboard-nine-grafico11-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'div-comparativo-marcas',
      nome,
    );
  }

  exportToPptImagemPura() {
    var nome = this.translate.instant('dashboard-nine-grafico11-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'imagem-pura',
      nome
    );
  }


  exportToPptEvolutivoMarcas() {
    var nome = this.translate.instant('dashboard-nine-grafico11-titulo');
    if (this.ModelAtributo) {
      nome = this.ModelAtributo.DescItem;
    }
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'grafico-linhas',
      nome
    );
  }


  exportToPptEvolutivoMarcas2() {
    var nome = this.translate.instant('dashboard-nine-grafico14-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'grafico-linhas2',
      nome
    );
  }

  exportToPptEvolutivoMarcas22() {
    var nome = this.translate.instant('dashboard-nine-grafico16-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'grafico-linhas22',
      nome
    );
  }


  exportToPptTabela() {
    var nome = this.translate.instant('dashboard-nine-grafico11-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'grafico-tabela',
      nome
    );
  }

  exportToPptTabelaBloco6() {
    var nome = this.translate.instant('dashboard-nine-grafico12-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'grafico-tabela-bloco-6',
      nome
    );
  }

  exportToPptTabelaBloco10() {
    var nome = this.translate.instant('dashboard-nine-grafico13-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'grafico-tabela-bloco-10',
      nome
    );
  }

  exportToPptTabelaBloco2() {
    var nome = this.translate.instant('dashboard-nine-grafico15-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'grafico-tabela-bloco-2',
      nome
    );
  }

}
