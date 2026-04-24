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
import { DashBoardFourService } from 'src/app/services/dashboard-four-service';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';
import { Session } from '../home/guards/session';
import { LogService } from 'src/app/services/log.service';



@Component({
  selector: 'app-dashboard-four',
  templateUrl: './dashboard-four.component.html',
  styleUrls: ['./dashboard-four.component.scss']
})
export class DashboardFourComponent implements OnInit {

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

  tipoBia1: boolean = true;
  tipoBia2: boolean = true;
  tipoBia3: boolean = true;

  ativaEvolutivoMarcasLinhas: boolean = true;

  paginaAtiva: boolean = true;

  constructor(public router: Router,
    public menuService: MenuService,
    public filtroService: FiltroGlobalService, private downloadArquivoService: DownloadArquivoService,
    private translate: TranslateService, private conversorPowerpointService: ConversorPowerpointService,
    private dashBoardFourService: DashBoardFourService,
    private session: Session,
    private logService: LogService,
  ) { }

  ngOnInit(): void {

    window.scroll({
      top: 0,
      left: 0,
      behavior: 'smooth'
    });



    this.menuService.nomePage = this.translate.instant('navbar.dashboard-four');
    this.ModelAtributo.IdItem = 1;

    this.menuService.activeMenu = 4;
    this.menuService.menuSelecao = "4"


    EventEmitterService.get("emit-dashboard-four").subscribe((x) => {
      this.paginaAtiva = false;
      this.marcasSelecionadasGrafico1 = new Array<PadraoComboFiltro>();
      this.marcasSelecionadasGrafico3 = new Array<PadraoComboFiltro>();
      this.carregarGraficos();
      this.logService.GravaLogRota(this.router.url).subscribe(
      );

    })
    // this.carregarEvolutivoMarcas();
  }

  carregarGraficos() {



    var idFiltro = this.filtroService.ModelTarget ? this.filtroService.ModelTarget.IdItem : 1

    this.filtroService.FiltroAtributos(idFiltro)
      .subscribe((response: Array<PadraoComboFiltro>) => {
        this.ModelAtributo = response[0];


        this.filtroService.FiltroMarcas(this.filtroService.ModelRegiao)
          .subscribe((response: Array<PadraoComboFiltro>) => {
            this.filtroService.listaMarcas = response;


            this.marcaColuna1 = null;
            this.marcaColuna2 = null;
            this.marcaColuna3 = null;
            this.marcaColuna4 = null;
            this.marcaColuna5 = null;

            this.carregaFiltroMarcas();
            this.ComparativoMarcas();
            this.carregarGraficosFirstLoad();
            this.carregarEvolutivoMarcas();

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
    filtros.Onda = new Array<PadraoComboFiltro>();
    filtros.Onda.push(this.filtroService.ModelOnda);
    filtros.CodUser = this.session.getCodUserSession();
    filtros.CodIdioma = 1;// this.authService.idDefaultLangUser;
    return filtros;
  }


  changetipoBia(nGrafico: number) {

    switch (nGrafico) {
      case 1:
        this.ComparativoMarcas();
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

  ComparativoMarcas() {

    var textoTooltipPerc = this.translate.instant('grafico-texto-tooltip-perc');
    var textoTooltipMedia = this.translate.instant('grafico-texto-tooltip-media');
    var textoTooltipBase = this.translate.instant('grafico-texto-tooltip-base');

    this.ativaComparativoMarcas = true;
    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.ParamBIA = this.tipoBia1 ? 1 : 2;

    if (!this.marcasSelecionadasGrafico1.length) {
      this.marcasSelecionadasGrafico1.push(this.filtroService.listaMarcas[0])
      this.marcasSelecionadasGrafico1.push(this.filtroService.listaMarcas[1])
      this.marcasSelecionadasGrafico1.push(this.filtroService.listaMarcas[2])
      this.marcasSelecionadasGrafico1.push(this.filtroService.listaMarcas[3])
    }


    filtros.Marca = this.marcasSelecionadasGrafico1;
    filtros.Sequencia = 1;

    this.dashBoardFourService.ComparativoMarcas(filtros)
      .subscribe((response: GraficoLinhasModel) => {

        this.ativaComparativoMarcas = true;

        this.graficoComparativoMarcasModel = response;

        this.verificaBaseComparativoMarcas();

        if (response == null) {
          this.ativaComparativoMarcas = false;
          return
        }

      }, (error) => console.error(error),
        () => {

          var grafico = [] as any;
          var index = 0;

          // var cores = this.getArrayColors(this.graficoComparativoMarcasModel.Grafico.length);
          var cores = this.graficoComparativoMarcasModel.Cores;

          if (this.graficoComparativoMarcasModel?.Grafico?.length) {

            // var lista = this.graficoComparativoMarcasModel.Grafico.slice(0, 10);
            var lista = this.graficoComparativoMarcasModel.Grafico;

            lista?.forEach(((item) => {

              // var dataArray = [];

              // item.data.forEach(x => {
              //   var novoItem = {
              //     "name": x.name,
              //     "y": x.y,
              //     "periodo": x.periodo,
              //     "valorbase": x.valorbase,
              //     // "segmentColor": 'black',
              //     // "color": 'black'
              //   };
              //   dataArray.push(novoItem);
              // });


              var objeto = {
                id: item.name,
                linkedTo: "legenda" + index,
                type: 'line',
                name: item.name,
                data: item.data,
                showInLegend: false,
                pointPlacement: 'on',

                marker: {
                  symbol: "circle",
                  fillColor: cores[index],
                  lineWidth: 0,
                  radius: 0,
                  lineColor: '#fff'
                },

                style: {
                  fontSize: '4px',
                  fontWeight: 'normal',
                  fontFamily: 'Poppins',
                  fontStyle: 'normal',
                  color: '#585656',
                  lineHeight: '20px',
                  textAlign: 'center',
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
                // dataLabels: {
                //   allowOverlap: true,
                //    enabled: true
                //  },
              }

              grafico.push(objeto);
              grafico.push(legenda);
              index++;
            }));

            this.graficoComparativoMarcasModel.Grafico = grafico;

            var graficoComparativoMarcasModelHighchart: any = {
              chart: {
                renderTo: 'comparativo-marcas',
                spacingTop: 15,
                // spacingBottom: 20,
                height: 650,
                // width:890,

                title: {
                  text: "Texto",
                  enabled: false,
                },
                labels: {
                  enabled: false,
                },
                type: "line",
                polar: true,
                // //backgroundColor: "#F8F8F8",
                style: {
                  fontFamily: "Poppins",
                },
              },
              title: {
                text: "",
                enabled: false,
              },
              pane: [{
                startAngle: -10,
                endAngle: 350,
                size: '90%',
              }],
              legend: {
                align: "left",
                verticalAlign: "bottom",
                enabled: true,
                alignColumns: true,
                itemDistance: 16,
                spacingBottom: 0,
                //spacingTop: 60,

                itemStyle: {
                  fontSize: '12px',
                  fontWeight: 'normal',
                  fontFamily: 'Poppins',
                  fontStyle: 'normal',
                  color: '#585656',
                  lineHeight: '18px',
                  textAlign: 'center',
                },

              },
              xAxis: {
                tickmarkPlacement: 'on',
                pane: 0,
                lineWidth: 0,
                startOnTick: true,
                lineColor: '#f3f3f3',
                categories: this.graficoComparativoMarcasModel.Periodos,
                tickInterval: 1,

                plotBands: [
                  {
                    color: "rgba(68, 170, 213, .2)",
                  },
                ],

                labels: {
                  useHTML: true,//Set to true
                  style: {
                    width: '300px',
                    whiteSpace: 'normal',//set to normal
                    fontSize: '10px',
                    fontWeight: '400',
                    fontFamily: 'Poppins',
                    fontStyle: 'normal',
                    color: '#3B3E44',
                    lineHeight: '8px',
                    textAlign: 'left',
                    height: '5px',
                  },
                  step: 1,
                  formatter: function () {//use formatter to break word.

                    // if (this.value == "I would pay more for") {
                    //   alert(this.pos)
                    //   return '<div align="" style="position:relative;top:0px;height:20px;display: flex;width:195px;margin: 0px;display: flex; align-items: center;">' + this.value + '</div>';
                    // }

                    if (this.pos == 0) {
                      return '<div align="" style="position:relative;top:-18px;height:0px;display: flex;width:195px;margin: 0px;display: flex; align-items: center;">' + this.value + '</div>';
                    }


                    if (this.pos == 14) {
                      return '<div align="" style="position:relative;top:1px;height:0px;display: flex;width:115px;margin: 0px;display: flex; align-items: center;">' + this.value + '</div>';
                    }

                    return '<div align="left" style="display: flex;word-wrap: break-word;width:170px">' + this.value + '</div>';
                  }

                },

                // labels: {
                //   enabled: true,
                //   floating: true,
                //    useHTML: true,
                //   allowOverlap: true,
                //    align: 'center',

                //   // formatter: function () {
                //   //   return '<div style="color: #6E778B !important;font-weight: 400px !important;font-size:12px !important;width: 150px;">' + this.value + '</div>';
                //   //   //   //
                //   //   //   if ((descPeriodo == this.value)) {
                //   //   //     return '<b style="color: black !important;font-weight: 600px !important;font-size:15px !important;">' + this.value + '</b>';
                //   //   //   }
                //   //   //   else {
                //   //   //     return this.value;
                //   //   //   }
                //   // },

                //   style: {
                //     fontSize: '12px',
                //     fontWeight: '400',
                //     fontFamily: 'Poppins',
                //     fontStyle: 'normal',
                //     color: '#3B3E44',
                //     lineHeight: '14px',
                //      textAlign: 'center',
                //     //  width: '150px',

                //      wordBreak: 'break-all',
                //     textOverflow: 'allow'
                //   },
                // },
              },
              yAxis: {
                min: 0,
                max: 110, // <<< define a escala máxima para 100
                tickInterval: 10,
                //  tickInterval: 20,
                gridLineInterpolation: 'polygon',
                pane: 0,
                lineWidth: 0,
                label: {
                  enabled: true,

                },
                title: {
                  text: "undef",
                  enabled: false,
                },
                labels: {
                  enabled: true,
                  align: 'left',
                  y: 5,
                  x: 10,

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


                  if (pointer?.baseminima != "") {

                    return '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px">  </div>  <div style="font-weight: 600;text-align: center; color:' + pointer?.color + '">   ' + pointer?.category + '</div>'
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 500;text-align: left;width: 95px"> ' + textoTooltipMedia + '  </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.media + '</div>'
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 500;text-align: left;width: 95px">  ' + textoTooltipPerc + ' </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.y + '%</div>  '
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 500;text-align: left;width: 95px"> ' + textoTooltipBase + '</div>  <div style="font-weight: 400;text-align: right;color:red;">   ' + pointer?.valorbase + '</div> ';
                  }
                  else {
                    return '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px">   </div>  <div style="font-weight: 600;text-align: center;color:' + pointer?.color + '">   ' + pointer?.category + '</div>'
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 500;text-align: left;width: 95px"> ' + textoTooltipMedia + '  </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.media + '</div>'
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 500;text-align: left;width: 95px">  ' + textoTooltipPerc + ' </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.y + '%</div>  '
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 500;text-align: left;width: 95px"> ' + textoTooltipBase + '</div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.valorbase + '</div> ';
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

                // column: {
                //     pointPadding: 0,
                //     groupPadding: 0
                // },
                areaspline: {
                  fillOpacity: 0.1,
                },
                // series: {
                //   //  lineWidth: 1,
                //   //  allowPointSelect: true
                // },

                line: {

                  dataLabels: {
                    enabled: false,
                    // formatter: function () {

                    //   if (this.y == 0) {
                    //     return '';
                    //   }

                    //   // if (this.y.toString().length < 3)
                    //   //   return "<div> <div>" + this.y.toString() + ',0' + "</div>" + '<br>' + "<div>" + this.x + "</div>" + "</div>";
                    //   // else
                    //   //   return "<div> <div>" + this.y.toString().replace(".", ",") + "</div>" + '<br>' + "<div>" + this.x + "</div>" + "</div>";

                    //   if (this.y.toString().length < 3)
                    //     return this.y.toString() + ',0';
                    //   else
                    //     return this.y.toString().replace(".", ",");


                    // },
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
              series: this.graficoComparativoMarcasModel.Grafico,
              colors: cores,
              exporting: {
                enabled: false,
              },
              // rules: [{
              //   condition: {
              //     maxWidth: 900
              //   },
              //   chartOptions: {
              //     legend: {
              //       layout: 'horizontal',
              //       align: 'center',
              //       verticalAlign: 'bottom',
              //       horizontalAlign: 'center',
              //     }
              //   }
              // }]
              responsive: {
                rules: [{
                  condition: {
                    maxWidth: 900,
                    // maxHeight: 200
                  },
                  chartOptions: {
                    // legend: {
                    //   align: 'center',
                    //   verticalAlign: 'bottom',
                    //   layout: 'horizontal'
                    // },
                    // pane: {
                    //   // size: '70%'
                    // }
                  }
                }]
              }
            }
          }





          if (this.graficoComparativoMarcasModel?.Grafico?.length) {
            Highcharts.chart(graficoComparativoMarcasModelHighchart).destroy();
            Highcharts.chart(graficoComparativoMarcasModelHighchart);
          }
          else {
            this.ativaComparativoMarcas = false;
          }

        }
      )


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
    // this.graficoLinha("container-coluna-5", 5, this.marcaColuna5);

  }




  listaGroup: ItemGrafico[];
  graficoLinha1() {

    var graficoImagemPuraEvolutivoColunaModel = new GraficoLinhasModel();

    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Marca.push(this.filtroService.listaMarcas[0]);
    filtros.Sequencia = 0;

    filtros.ParamBIA = this.tipoBia2 ? 1 : 2;

    this.dashBoardFourService.ImagemEvolutiva(filtros)
      .subscribe((response: GraficoLinhasModel) => {

        this.graficoImagemPuraEvolutivoColuna1Model = response;

      }, (error) => console.error(error),
        () => {

          var grafico = [] as any;
          var index = 0;
          var cores = ['transparent'];

          if (this.graficoImagemPuraEvolutivoColuna1Model?.Grafico?.length) {

            var lista = this.graficoImagemPuraEvolutivoColuna1Model.Grafico.reverse();

            // ── Extrai dados de grupo da primeira série (todas têm os mesmos atributos) ──
            const primeiraSerieData: any[] = (lista[0]?.data || []);

            // Monta plotBands: faixa colorida para cada grupo de atributos consecutivos
            const plotBands: any[] = [];
            const gruposLegenda: { descricao: string; cor: string }[] = [];
            let grupoAtual = '';
            let bandStart = 0;


            const mapa = new Map<string, any>();

            primeiraSerieData.forEach((ponto, i) => {
              const grupo = ponto.GrupoAtributo || '';

              if (!mapa.has(grupo)) {
                mapa.set(grupo, {
                  index: i,
                  grupo,
                  cor: ponto.CorAtributo || '#ffffff'
                });
              }
            });

            this.listaGroup = Array.from(mapa.values());


            primeiraSerieData.forEach((ponto: any, i: number) => {
              const grupo = ponto.GrupoAtributo || '';
              const cor = ponto.CorAtributo || '#ffffff';

              if (grupo !== grupoAtual) {
                // Fecha banda anterior
                if (i > 0) {
                  plotBands.push({
                    from: bandStart - 0.5,
                    to: i - 0.5,
                    color: this._hexToRgba(gruposLegenda.length > 0
                      ? gruposLegenda[gruposLegenda.length - 1].cor
                      : '#ffffff', 0.15),
                    zIndex: 0
                  });
                }
                // Registra legenda de grupo (sem duplicar)
                if (!gruposLegenda.find(g => g.descricao === grupo)) {
                  gruposLegenda.push({ descricao: grupo, cor });
                }
                grupoAtual = grupo;
                bandStart = i;
              }
            });

            // Fecha última banda
            if (primeiraSerieData.length > 0) {
              plotBands.push({
                from: bandStart - 0.5,
                to: primeiraSerieData.length - 0.5,
                color: this._hexToRgba(gruposLegenda.length > 0
                  ? gruposLegenda[gruposLegenda.length - 1].cor
                  : '#ffffff', 0.15),
                zIndex: 0
              });
            }

            // ── Renderiza legenda de grupos abaixo do container do gráfico ──
            this._renderizaLegendaGrupos('container-coluna-0-legenda', gruposLegenda);

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
                spacingLeft: 18,

                height: 858,
                width: 400,
                title: {
                  text: "Texto",
                  enabled: false,
                },
                labels: {
                  enabled: true,
                },
                type: "spline",
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
                  fontSize: '12px',
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

                categories: this.graficoImagemPuraEvolutivoColuna1Model.Periodos,
                plotBands: plotBands,   // ← cores de fundo por grupo
                labels: {
                  enabled: true,
                  html: true,
                  floating: true,
                  align: 'left',
                  x: 10, // move tudo pra direita sem quebrar o layout

                  //                   formatter: function () {
                  //   const index = this.pos;
                  //   const chart = this.chart;
                  //   const serie = chart.series[0]; // ou ajuste se tiver várias

                  //   const value = serie.yData[index];

                  //   // return `
                  //   //   <div style="margin-left:40px; font-size:12px; font-family:Poppins;">
                  //   //     ${this.value} - <b>${value}</b>
                  //   //   </div>
                  //   // `;
                  //   return `
                  //     <div  style="width: 10px; margin-left:40px; font-size:12px; font-family:Poppins; position:relative; right:30px;">
                  //     <div style="position:absolute; right:30px;">  ${this.value}     </div>
                  //     </div>
                  //   `;
                  // },

                  style: {
                    fontSize: '12px',
                    fontWeight: 'normal',
                    fontFamily: 'Poppins',
                    fontStyle: 'normal',
                    color: '#585656',
                    lineHeight: '20px',
                    textAlign: 'left',
                    width: 300,

                  },
                },
              }, {
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
                label: { enabled: true },
                title: { text: "undef", enabled: false },
                labels: {
                  enabled: false,
                  align: 'right',
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
                    + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 95px">  Base </div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.valorbase + '</div>  ';
                },
              },
              credits: { enabled: false },
              plotOptions: {
                areaspline: { fillOpacity: 0.1 },
                enableMouseTracking: true,
                line: {
                  fillColor: '#FFFFFF',
                  dataLabels: {
                    enabled: false,
                    useHTML: true,
                    formatter: function () {
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
              exporting: { enabled: false },
              rules: [{
                condition: { maxWidth: 1000 },
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

            if (this.graficoImagemPuraEvolutivoColuna1Model?.Grafico?.length) {
              Highcharts.chart(graficoImagemPuraEvolutivoColuna1ModelHighchart).destroy();
              Highcharts.chart(graficoImagemPuraEvolutivoColuna1ModelHighchart);
            }
          }
        }
      )
  }

  // ── Converte hex em rgba com opacidade ──────────────────────────────────────
  private _hexToRgba(hex: string, alpha: number): string {
    if (!hex || hex === '') return `rgba(200,200,200,${alpha})`;
    const clean = hex.replace('#', '');
    const r = parseInt(clean.substring(0, 2), 16);
    const g = parseInt(clean.substring(2, 4), 16);
    const b = parseInt(clean.substring(4, 6), 16);
    return `rgba(${r},${g},${b},${alpha})`;
  }

  // ── Renderiza legenda de grupos como HTML abaixo do gráfico ────────────────
  private _renderizaLegendaGrupos(
    containerId: string,
    grupos: { descricao: string; cor: string }[]
  ): void {
    // Aguarda o DOM estar pronto
    setTimeout(() => {
      const el = document.getElementById(containerId);
      if (!el) return;

      el.innerHTML = '';
      el.style.display = 'flex';
      el.style.flexWrap = 'wrap';
      el.style.gap = '8px';
      el.style.padding = '8px 0';
      el.style.fontFamily = 'Poppins';

      grupos.forEach(grupo => {
        const item = document.createElement('div');
        item.style.display = 'flex';
        item.style.alignItems = 'center';
        item.style.gap = '6px';
        item.style.fontSize = '12px';
        item.style.color = '#585656';

        const bolinha = document.createElement('div');
        bolinha.style.width = '14px';
        bolinha.style.height = '14px';
        bolinha.style.borderRadius = '3px';
        bolinha.style.backgroundColor = grupo.cor || '#cccccc';
        bolinha.style.flexShrink = '0';

        const texto = document.createElement('span');
        texto.textContent = grupo.descricao;

        item.appendChild(bolinha);
        item.appendChild(texto);
        el.appendChild(item);
      });
    }, 100);
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
    filtros.ParamBIA = this.tipoBia2 ? 1 : 2;

    this.dashBoardFourService.ImagemEvolutiva(filtros)
      .subscribe((response: GraficoLinhasModel) => {

        graficoImagemPuraEvolutivoColunaModel = response;

        this.verificaBaseGraficoLinha(graficoImagemPuraEvolutivoColunaModel);

      }, (error) => console.error(error),
        () => {

          var grafico = [] as any;
          var index = 0;

          // var cores = this.getArrayColors(graficoImagemPuraEvolutivoColunaModel.Grafico.length);
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
                height: 858,
                width: 120,
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
                    align: "right",
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


  carregarEvolutivoMarcas() {

    var textoTooltipPerc = this.translate.instant('grafico-texto-tooltip-perc');
    var textoTooltipMedia = this.translate.instant('grafico-texto-tooltip-media');
    var textoTooltipBase = this.translate.instant('grafico-texto-tooltip-base');
    var textoTooltipPeriodo = this.translate.instant('grafico-texto-tooltip-periodo');

    var graficoEvolutivoMarcasLinhaModel = new GraficoLinhasModel();

    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Atributo = (this.ModelAtributo) ? this.ModelAtributo.IdItem : 1;
    filtros.ParamBIA = this.tipoBia3 ? 1 : 2;

    if (!this.marcasSelecionadasGrafico3.length) {
      this.marcasSelecionadasGrafico3.push(this.filtroService.listaMarcas[0])
      this.marcasSelecionadasGrafico3.push(this.filtroService.listaMarcas[1])
      this.marcasSelecionadasGrafico3.push(this.filtroService.listaMarcas[2])
      this.marcasSelecionadasGrafico3.push(this.filtroService.listaMarcas[3])

    }

    filtros.Marca = this.marcasSelecionadasGrafico3;
    filtros.Sequencia = 6;

    this.dashBoardFourService.ImagemEvolutivaLinhas(filtros)
      .subscribe((response: GraficoLinhasModel) => {

        this.ativaEvolutivoMarcasLinhas = true;
        graficoEvolutivoMarcasLinhaModel = response;

        this.verificaBaseEvolutivoMarcas(graficoEvolutivoMarcasLinhaModel);

      }, (error) => console.error(error),
        () => {

          var grafico = [] as any;
          var index = 0;
          // var cores = this.getArrayColors(graficoEvolutivoMarcasLinhaModel.Grafico.length);
          var cores = graficoEvolutivoMarcasLinhaModel.Cores;

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
                height: 350,
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
                width: 290,
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
                      + '<div style="max-height: 0px !important;font-family: Poppins;font-style: normal;font-weight: 600;text-align: left;width: 145px;"> ' + textoTooltipPeriodo + '</div>  <div style="font-weight: 400;text-align: right;">   ' + pointer?.periodo + '</div> ';
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

          if (graficoEvolutivoMarcasLinhaModel?.Grafico?.length) {
            Highcharts.chart(graficoEvolutivoMarcasLinhaModelHighchart).destroy();
            Highcharts.chart(graficoEvolutivoMarcasLinhaModelHighchart);
          }

        }
      )
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

  onChangeListaMarcas(marcas: Array<PadraoComboFiltro>) {

    this.marcasSelecionadasGrafico1 = marcas

    if (!this.marcasSelecionadasGrafico1.length) {
      this.marcasSelecionadasGrafico1.push(this.filtroService.listaMarcas[0])
      this.marcasSelecionadasGrafico1.push(this.filtroService.listaMarcas[1])
      this.marcasSelecionadasGrafico1.push(this.filtroService.listaMarcas[2])
      this.marcasSelecionadasGrafico1.push(this.filtroService.listaMarcas[3])
    }
    this.ComparativoMarcas();
  }


  onChangeListaMarcasGrafico3(marcas: Array<PadraoComboFiltro>) {
    this.marcasSelecionadasGrafico3 = marcas

    if (!this.marcasSelecionadasGrafico3.length) {
      this.marcasSelecionadasGrafico3.push(this.filtroService.listaMarcas[0])
      this.marcasSelecionadasGrafico3.push(this.filtroService.listaMarcas[1])
      this.marcasSelecionadasGrafico3.push(this.filtroService.listaMarcas[2])
      this.marcasSelecionadasGrafico3.push(this.filtroService.listaMarcas[3])
    }
    this.carregarEvolutivoMarcas();
  }


  onChangeAtributo(atributo: PadraoComboFiltro) {
    this.ModelAtributo = atributo;
    this.carregarEvolutivoMarcas();
  }

  ajustaScrollFinal() {
    if (document.getElementById('aqui_id_div_grafico') != null && !document.getElementById('aqui_id_div_grafico').scrollLeft) {
      document.getElementById('aqui_id_div_grafico').scrollLeft = 9999;
    }

  }

  async exportToExcelGrafico1() {
    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Marca = this.marcasSelecionadasGrafico1;
    filtros.Sequencia = 1;
    filtros.ParamBIA = this.tipoBia1 ? 1 : 2;
    var titulo = this.translate.instant('dashboard-four-grafico1-titulo');
    filtros.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardFourComparativoMarcas(filtros)
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
    var titulo = this.translate.instant('dashboard-four-grafico2-titulo');
    filtroPadrao.TituloGrafico = titulo;
    filtroPadrao.ParamBIA = this.tipoBia2 ? 1 : 2;
    await this.downloadArquivoService.
      DownloadDashboardFourEvolutivoPeriodos(filtroPadrao)
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
    var filtros = this.carregaFiltros();
    filtros.Marca = new Array<PadraoComboFiltro>();
    filtros.Marca = this.marcasSelecionadasGrafico3;
    filtros.Sequencia = 1;
    filtros.Atributo = this.ModelAtributo.IdItem;
    filtros.ParamBIA = this.tipoBia3 ? 1 : 2;

    var titulo = this.translate.instant('dashboard-four-grafico3-titulo');
    filtros.TituloGrafico = titulo;
    await this.downloadArquivoService.
      DownloadDashboardFourComparativoImagemPura(filtros)
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
    var nome = this.translate.instant('dashboard-four-grafico1-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'div-comparativo-marcas',
      nome,
    );
  }

  exportToPptImagemPura() {
    var nome = this.translate.instant('dashboard-four-grafico2-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'imagem-pura',
      nome
    );
  }


  exportToPptEvolutivoMarcas() {
    var nome = this.translate.instant('dashboard-four-grafico3-titulo');
    this.conversorPowerpointService.captureScreenPPTAlternative(
      'grafico-linhas',
      nome
    );
  }

}


export class ItemGrafico {
  index: number;
  grupo: string;
  cor: string;
}
