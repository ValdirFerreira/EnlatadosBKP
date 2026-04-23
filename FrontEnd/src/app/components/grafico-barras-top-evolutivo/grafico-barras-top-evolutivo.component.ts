import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { GraficoColunaModel } from 'src/app/models/grafico-coluna/grafico-coluna';
import { GraficoBVCEvolutivo } from 'src/app/models/GraficoBVCEvolutivo/GraficoBVCEvolutivo';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';


@Component({
    selector: 'grafico-barras-top-evolutivo',
    templateUrl: './grafico-barras-top-evolutivo.component.html',
    styleUrls: ['./grafico-barras-top-evolutivo.component.scss']
})
export class GraficoBarrasTopEvolutivoComponent implements OnInit {

    @Input() filtro: PadraoComboFiltro;


    @Input() graficoBVCEvolutivo1: Array<GraficoBVCEvolutivo>;
    @Input() graficoBVCEvolutivo2: Array<GraficoBVCEvolutivo>;
    @Input() graficoBVCEvolutivo3: Array<GraficoBVCEvolutivo>;
    @Input() graficoBVCEvolutivo4: Array<GraficoBVCEvolutivo>;
    @Input() graficoBVCEvolutivo5: Array<GraficoBVCEvolutivo>;


    @Output('ondaColuna11') ondaColuna11: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna22') ondaColuna22: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna33') ondaColuna33: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna44') ondaColuna44: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna55') ondaColuna55: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

    onda1: PadraoComboFiltro;
    onda2: PadraoComboFiltro;
    onda3: PadraoComboFiltro;
    onda4: PadraoComboFiltro;


    qtdLinhas: number = 0;
    constructor(
        private translate: TranslateService, public filtroService: FiltroGlobalService,
    ) { }


    ngOnInit(): void {
        this.FiltroOndas();
    }

    ngOnChanges(changes: any) {

        if (this.graficoBVCEvolutivo1.length > 0) {
            this.qtdLinhas = this.graficoBVCEvolutivo1.length - 1;
        }

        if (this.graficoBVCEvolutivo2.length > 0) {
            this.qtdLinhas = this.graficoBVCEvolutivo2.length - 1;
        }

        if (this.graficoBVCEvolutivo3.length > 0) {
            this.qtdLinhas = this.graficoBVCEvolutivo3.length - 1;
        }

        if (this.graficoBVCEvolutivo4.length > 0) {
            this.qtdLinhas = this.graficoBVCEvolutivo4.length - 1;
        }

    }


    FiltroOndas() {

        if (this.filtroService.listaOnda) {
            var listaOndas = this.filtroService.listaOnda;

            this.onda1 = listaOndas[3];
            this.onda2 = listaOndas[2];
            this.onda3 = listaOndas[1];
            this.onda4 = listaOndas[0];
        }
        else {

            this.filtroService.FiltroOnda()
                .subscribe((response: Array<PadraoComboFiltro>) => {
                    var listaOndas = response;

                    this.onda1 = listaOndas[3];
                    this.onda2 = listaOndas[2];
                    this.onda3 = listaOndas[1];
                    this.onda4 = listaOndas[0];

                }, (error) => console.error(error),
                    () => {
                    }
                )
        }
    }

    porcBarraPositiva(valor) {
        if (valor > 0 && valor < 1)
        return "2%";

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


    montaCorBarraBackground(item: GraficoBVCEvolutivo, colPerc: number) {

        var corPositiva = "#F2994A";
        var corNegativa = "#F2994A";


        switch (this.filtro.IdItem) {
            case 0:
            case 1:
                corPositiva = "#F2994A";
                corNegativa = "#F2994A";
                break;

            case 2:
                corPositiva = "#27AE60";
                corNegativa = "#EB5757";
                break;

            case 3:
                corPositiva = "#4056B7";
                corNegativa = "#4056B7";
                break;

            default:
                corPositiva = "#F2994A";
                corNegativa = "#F2994A";
                break;
        }

        var cor = corPositiva;

        if (colPerc == 2) {
            if (item.Perc1 < 0)
                cor = corNegativa;
        }

        return 'border:1px solid ' + cor + ' !important;';

    }


    

    montaCorBarra(item: GraficoBVCEvolutivo, colPerc: number) {



        var corPositiva = "#F2994A";
        var corNegativa = "#F2994A";


        switch (this.filtro.IdItem) {
            case 0:
            case 1:
                corPositiva = "#F2994A";
                corNegativa = "#F2994A";
                break;

            case 2:
                corPositiva = "#27AE60";
                corNegativa = "#EB5757";
                break;

            case 3:
                corPositiva = "#4056B7";
                corNegativa = "#4056B7";
                break;

            default:
                corPositiva = "#F2994A";
                corNegativa = "#F2994A";
                break;
        }

        var cor = corPositiva;

        // if (colPerc == 2) {
        //     if (item.Perc1 < 0)
        //         cor = corNegativa;
        // }

        if (item.Perc1 < 0)
            cor = corNegativa;

  
          //  return '1px solid red' + cor + ' !important;';

         return '1px solid '+cor;
    

    }

    montaCorBarraDinamica() {

        var corPositiva = "#F2994A";
        var corNegativa = "#F2994A";


        switch (this.filtro.IdItem) {
            case 0:
            case 1:
                corPositiva = "#F2994A";
                corNegativa = "#F2994A";
                break;

            case 2:
                corPositiva = "#27AE60";
                corNegativa = "#EB5757";
                break;

            case 3:
                corPositiva = "#4056B7";
                corNegativa = "#4056B7";
                break;

            default:
                corPositiva = "#F2994A";
                corNegativa = "#F2994A";
                break;
        }

        var cor = corPositiva;



        return cor;

    }



    onchangeOndaColuna1(item: PadraoComboFiltro) {
        this.ondaColuna11.emit(item);
    }
    onchangeOndaColuna2(item: PadraoComboFiltro) {
        this.ondaColuna22.emit(item);
    }
    onchangeOndaColuna3(item: PadraoComboFiltro) {
        this.ondaColuna33.emit(item);
    }
    onchangeOndaColuna4(item: PadraoComboFiltro) {
        this.ondaColuna44.emit(item);
    }
    onchangeOndaColuna5(item: PadraoComboFiltro) {
        this.ondaColuna55.emit(item);
    }


    verificaMsgBase() {

        if (this.graficoBVCEvolutivo1.length && this.graficoBVCEvolutivo1?.[0].BaseMinimaPerc1) {
            return this.graficoBVCEvolutivo1[0].BaseMinimaPerc1;
        }

        if (this.graficoBVCEvolutivo2.length && this.graficoBVCEvolutivo2?.[0].BaseMinimaPerc1) {
            return this.graficoBVCEvolutivo2[0].BaseMinimaPerc1;
        }

        if (this.graficoBVCEvolutivo3.length && this.graficoBVCEvolutivo3?.[0].BaseMinimaPerc1) {
            return this.graficoBVCEvolutivo3[0].BaseMinimaPerc1;
        }

        if (this.graficoBVCEvolutivo4.length && this.graficoBVCEvolutivo4?.[0].BaseMinimaPerc1) {
            return this.graficoBVCEvolutivo4[0].BaseMinimaPerc1;
        }


        return '';
    }

}