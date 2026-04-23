import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';
import { GraficoColunaModel } from 'src/app/models/grafico-coluna/grafico-coluna';
import { GraficoImagemPosicionamento } from 'src/app/models/grafico-Imagem-posicionamento/GraficoImagemPosicionamento';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';


@Component({
    selector: 'grafico-barras-marcas-triplo',
    templateUrl: './grafico-barras-marcas-triplo.component.html',
    styleUrls: ['./grafico-barras-marcas-triplo.component.scss']
})
export class GraficoBarrasMarcasTriploComponent implements OnInit {

    @Input() graficoColunaModel1: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel2: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel1_3: Array<GraficoImagemPosicionamento>;

    @Input() graficoColunaModel3: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel4: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel3_3: Array<GraficoImagemPosicionamento>;

    @Input() graficoColunaModel5: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel6: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel5_3: Array<GraficoImagemPosicionamento>;

    @Input() graficoColunaModel7: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel8: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel7_3: Array<GraficoImagemPosicionamento>;

    @Input() graficoColunaModel9: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel10: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel9_3: Array<GraficoImagemPosicionamento>;

    @Output('marcaColuna1') marcaColuna1: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna2') marcaColuna2: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna3') marcaColuna3: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna4') marcaColuna4: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna5') marcaColuna5: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();


    @Output('ondaColuna1') ondaColuna1: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna2') ondaColuna2: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna1_3') ondaColuna1_3: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();


    @Output('ondaColuna3') ondaColuna3: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna4') ondaColuna4: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna3_3') ondaColuna3_3: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

    @Output('ondaColuna5') ondaColuna5: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna6') ondaColuna6: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna5_3') ondaColuna5_3: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

    @Output('ondaColuna7') ondaColuna7: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna8') ondaColuna8: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna7_3') ondaColuna7_3: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

    @Output('ondaColuna9') ondaColuna9: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna10') ondaColuna10: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna9_3') ondaColuna9_3: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();


    onda1: PadraoComboFiltro;
    onda2: PadraoComboFiltro;
    onda1_3: PadraoComboFiltro;

    onda3: PadraoComboFiltro;
    onda4: PadraoComboFiltro;
    onda3_3: PadraoComboFiltro;

    onda5: PadraoComboFiltro;
    onda6: PadraoComboFiltro;
    onda5_3: PadraoComboFiltro;

    onda7: PadraoComboFiltro;
    onda8: PadraoComboFiltro;
    onda7_3: PadraoComboFiltro;

    onda9: PadraoComboFiltro;
    onda10: PadraoComboFiltro;
    onda9_3: PadraoComboFiltro;

    qtdLinhas: number = 0;



    ngOnInit(): void {
        this.FiltroOndas();
        this.qtdLinhas = this.graficoColunaModel1.length - 1;
    }

    constructor(
        public filtroService: FiltroGlobalService,
    ) { }


    porcBarraPositiva(valor) {
        if (valor <= 0)
            return "0%";
        return valor + "%";
    }

    porcBarraNegativa(valor) {
        if (valor >= 0)
            return "0%";

        let nValor = valor * -1;
        return nValor + "%";
    }


    onchangeMarcaColuna1(item: PadraoComboFiltro) {
        this.marcaColuna1.emit(item);
    }
    onchangeMarcaColuna2(item: PadraoComboFiltro) {
        this.marcaColuna2.emit(item);
    }
    onchangeMarcaColuna3(item: PadraoComboFiltro) {
        this.marcaColuna3.emit(item);
    }
    onchangeMarcaColuna4(item: PadraoComboFiltro) {
        this.marcaColuna4.emit(item);
    }
    onchangeMarcaColuna5(item: PadraoComboFiltro) {
        this.marcaColuna5.emit(item);
    }



    onchangeOndaColuna1(item: PadraoComboFiltro) {
        this.ondaColuna1.emit(item);
    }
    onchangeOndaColuna2(item: PadraoComboFiltro) {
        this.ondaColuna2.emit(item);
    }
    onchangeOndaColuna1_3(item: PadraoComboFiltro) {
        this.ondaColuna1_3.emit(item);
    }


    onchangeOndaColuna3(item: PadraoComboFiltro) {
        this.ondaColuna3.emit(item);
    }    
    onchangeOndaColuna4(item: PadraoComboFiltro) {
        this.ondaColuna4.emit(item);
    }
    onchangeOndaColuna3_3(item: PadraoComboFiltro) {
        this.ondaColuna3_3.emit(item);
    }   


    onchangeOndaColuna5(item: PadraoComboFiltro) {
        this.ondaColuna5.emit(item);
    }
    onchangeOndaColuna6(item: PadraoComboFiltro) {
        this.ondaColuna6.emit(item);
    }
    onchangeOndaColuna5_3(item: PadraoComboFiltro) {
        this.ondaColuna5_3.emit(item);
    }


    onchangeOndaColuna7(item: PadraoComboFiltro) {
        this.ondaColuna7.emit(item);
    }
    onchangeOndaColuna8(item: PadraoComboFiltro) {
        this.ondaColuna8.emit(item);
    }
    onchangeOndaColuna7_3(item: PadraoComboFiltro) {
        this.ondaColuna7_3.emit(item);
    }

    onchangeOndaColuna9(item: PadraoComboFiltro) {
        this.ondaColuna9.emit(item);
    }
    onchangeOndaColuna10(item: PadraoComboFiltro) {
        this.ondaColuna10.emit(item);
    }
    onchangeOndaColuna9_3(item: PadraoComboFiltro) {
        this.ondaColuna9_3.emit(item);
    }


    FiltroOndas() {

        if (this.filtroService.listaOnda) {
            var listaOndas = this.filtroService.listaOnda;

            this.onda1 = listaOndas[2];
            this.onda2 = listaOndas[1];
            this.onda1_3 = listaOndas[0];

            this.onda3 = listaOndas[2];
            this.onda4 = listaOndas[1];
            this.onda3_3 = listaOndas[0];

            this.onda5 = listaOndas[2];
            this.onda6 = listaOndas[1];            
            this.onda5_3 = listaOndas[0];

            this.onda7 = listaOndas[2];
            this.onda8 = listaOndas[1];
            this.onda7_3 = listaOndas[0];

            this.onda9 = listaOndas[2];
            this.onda10 = listaOndas[1];
            this.onda9_3 = listaOndas[0];
        }
        else {

            this.filtroService.FiltroOnda()
                .subscribe((response: Array<PadraoComboFiltro>) => {
                    var listaOndas = response;

                    this.onda1 = listaOndas[2];
                    this.onda2 = listaOndas[1];
                    this.onda1_3 = listaOndas[0];
        
                    this.onda3 = listaOndas[2];
                    this.onda4 = listaOndas[1];
                    this.onda3_3 = listaOndas[0];
        
                    this.onda5 = listaOndas[2];
                    this.onda6 = listaOndas[1];            
                    this.onda5_3 = listaOndas[0];
        
                    this.onda7 = listaOndas[2];
                    this.onda8 = listaOndas[1];
                    this.onda7_3 = listaOndas[0];
        
                    this.onda9 = listaOndas[2];
                    this.onda10 = listaOndas[1];
                    this.onda9_3 = listaOndas[0];

                }, (error) => console.error(error),
                    () => {
                    }
                )
        }
    }

    montaCorBarra(item: GraficoImagemPosicionamento) {
        return 'border:1px solid ' + item.CorSiteMarca + ' !important;';

    }


    montaCorBarraTesteSig(item: GraficoImagemPosicionamento) {
        if (item.PercAtual <= -21 || item.PercAtual > 20)
            return 'border:1px solid ' + item.CorSiteMarca + ' !important;';
        else
            return '';

    }


    verificaMsgBase() {
 
  const sufixos = [
    '',
    '_3'
  ];

  // Lista de índices de gráficos
  const indices = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];

  for (const i of indices) {
    for (const suf of sufixos) {
      const modelo = this[`graficoColunaModel${i}${suf}`];
      const baseMinima = modelo?.[0]?.BaseMinima;
      if (baseMinima) {
        return baseMinima;
      }
    }
  }

  return '';
}



    

}