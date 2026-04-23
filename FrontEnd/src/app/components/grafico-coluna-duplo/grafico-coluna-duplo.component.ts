import { Component, Input, OnInit } from '@angular/core';
import { Router } from '@angular/router'
import { GraficoColunaDuplo } from 'src/app/models/GraficoColunaDuplo/GraficoColunaDuplo';
import { GraficoComunicacaoRecall } from 'src/app/models/GraficoComunicacaoRecall/GraficoComunicacaoRecall';
import { GraficoComunicacaoVisto } from 'src/app/models/GraficoComunicacaoVisto/GraficoComunicacaoVisto';



@Component({
  selector: 'grafico-coluna-duplo',
  templateUrl: './grafico-coluna-duplo.component.html',
  styleUrls: ['./grafico-coluna-duplo.component.scss']
})
export class GraficoColunaDuploComponent implements OnInit {
  // @Input() graficoSatisfacaoGeral: Array<DadosTratadosGraficoSatisfacaoGeral>;


  @Input() graficoSatisfacaoGeralModel: Array<GraficoColunaDuplo>;
  @Input() graficoComunicacaoRecall: Array<GraficoComunicacaoRecall>;

  @Input() corGrafico: number;
  // 1 Amarela
  // 2 Azul
  @Input() graficoHoras: boolean = true;

  maxQtd = 0;


  containerErro: boolean;

  constructor(
    public router: Router,
  ) { }

  ngOnInit(): void {

  }


  converterFloatUmaCasaDecimal(valor): number {
    // return valor.toFixed(1).replace(".", ",");
    return valor;
  }

  calcularPorcentagemColuna(valor): any {

    if (valor > 100)
      return 100;

    return valor;
  }


  calcularPorcentagem(valor): number {
    return valor / 10;
  }

  calcularPorcentagemPontilhado(valor): number {
    return 100 - (valor / 10);
  }


  validaSig(valor: string) {

    if (valor == 'MAIOR')
      return 'sig-positive';

    if (valor == 'MENOR')
      return 'sig-negative';

    if (valor == 'IGUAL')
      return 'sig-vazio';

    return 'sig-vazio';
  }

  verificaMsgBase() {

    if (this.validaObjeto(this.graficoComunicacaoRecall)) {
      return this.graficoComunicacaoRecall?.[0].VisibilidadeBaseMinima;
    }

    if (this.validaObjeto(this.graficoComunicacaoRecall)) {
      return this.graficoComunicacaoRecall?.[0].LinkageBaseMinima;
    }

    if (this.validaObjeto(this.graficoComunicacaoRecall)) {
      return this.graficoComunicacaoRecall?.[0].RecalBaseMinima;
    }
    return '';
  }


  validaObjeto(obj: Array<GraficoComunicacaoRecall>) {
    if (obj != null && obj != undefined && obj.length > 0) {
        return true;
    }

    return false;
}

}

