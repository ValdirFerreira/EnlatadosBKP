import { Component, Input, OnInit } from '@angular/core';
import { Router } from '@angular/router'
import { GraficoColunaDuplo } from 'src/app/models/GraficoColunaDuplo/GraficoColunaDuplo';
import { GraficoComunicacaoVisto } from 'src/app/models/GraficoComunicacaoVisto/GraficoComunicacaoVisto';



@Component({
  selector: 'grafico-colunas-simples',
  templateUrl: './grafico-colunas-simples.component.html',
  styleUrls: ['./grafico-colunas-simples.component.scss']
})
export class GraficoColunasSimplesComponent implements OnInit {
  // @Input() graficoSatisfacaoGeral: Array<DadosTratadosGraficoSatisfacaoGeral>;


  @Input() graficoSatisfacaoGeralModel: Array<GraficoColunaDuplo>;

  @Input() graficoComunicacaoVisto = new Array<GraficoComunicacaoVisto>();

  @Input() corGrafico: number;
  // 1 Amarela
  // 2 Azul
  @Input() graficoHoras: boolean = false;

  maxQtd = 0;


  containerErro: boolean;

  constructor(
    public router: Router,
  ) { }

  ngOnInit(): void {


  }

  carregaGrafico() {

  }


  detalhesNPS(GraficoColunaModel: GraficoColunaDuplo) {
    // this.PrincipalFiltroVariaveisService.Periodo = GraficoColunaModel.CodPeriodo;
    // this.router.navigate(['/drilldown-satisfacao-detalhes']) 
  }

  converterFloatUmaCasaDecimal(valor): number {
    // return valor.toFixed(1).replace(".", ",");
    return valor;
  }

  calcularPorcentagemColuna(valor): any {

    return valor;

    if (this.maxQtd > 0) {

      var retorno = ((valor / this.maxQtd) * 100).toFixed(8);

      var retInt = parseInt(retorno)

      if (retInt < 1)
        retInt = valor / this.maxQtd;

      return retInt;
    }

  }

  calcularPorcentagem(valor): number {
    return valor / 10;
  }

  calcularPorcentagemPontilhado(valor): number {
    return 100 - (valor / 10);
  }

  validaSig(graficoColunaDuplo: GraficoColunaDuplo) {


    if (graficoColunaDuplo.Sig == 'MAIOR')
      return 'sig-positive';

    if (graficoColunaDuplo.Sig == 'MENOR')
      return 'sig-negative';

    if (graficoColunaDuplo.Sig == 'IGUAL')
      return 'sig-vazio';

    return 'sig-vazio';
  }

  verificaMsgBase() {

    var baseMinima = '';

    this.graficoComunicacaoVisto.forEach(item => {
      if (item.BaseMinima != '')
        baseMinima = item.BaseMinima;
    })

    return baseMinima;
  }

}

