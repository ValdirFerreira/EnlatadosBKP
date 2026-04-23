import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { GraficoColunaModel } from 'src/app/models/grafico-coluna/grafico-coluna';
import { GraficoBVCTop10 } from 'src/app/models/GraficoBVCTop10/GraficoBVCTop10';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';


@Component({
    selector: 'grafico-barras-top-marcas',
    templateUrl: './grafico-barras-top-marcas.component.html',
    styleUrls: ['./grafico-barras-top-marcas.component.scss']
})
export class GraficoBarrasTopMarcasComponent implements OnInit {

    @Input() graficoBVCTop10: Array<GraficoBVCTop10>;

    constructor(
        private translate: TranslateService,
    ) { }

    qtdLinhas: number = 0;

    ngOnInit(): void {


    }

    ngOnChanges(changes: any) {

        this.qtdLinhas = this.graficoBVCTop10.length - 1;
    }

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


    montaCorBarraEfeitos(item: GraficoBVCTop10) {

        var corPositiva = "#27AE60";
        var corNegativa = "#EB5757";

        var cor = corPositiva;

        if (item.Efeitos < 0)
            cor = corNegativa;

        return 'border:1px solid ' + cor + ' !important;';

    }

    montaCorBarraShareEfeitos(item: GraficoBVCTop10) {
        var corPositiva = "#27AE60";
        var corNegativa = "#EB5757";

        var cor = corPositiva;

        if (item.Efeitos < 0)
            cor = corNegativa;

        // if (item.Efeitos < -21 || item.Efeitos > 20)
        //     return 'border:1px solid ' + cor + ' !important;';
        // else
            return '';

    }



    montaCorBarraEquity(item: GraficoBVCTop10) {

        var cor = '#4056B7';
        return 'border:0px solid ' + cor + ' !important;';

    }

    montaCorBarraShareEquity(item: GraficoBVCTop10) {

        var cor = '#4056B7';


        // if (item.Efeitos < -21 || item.Efeitos > 20)
        //     return 'border:1px solid ' + cor + ' !important;';
        // else
            return '';

    }


    verificaMsgBase() {
        if (!this.graficoBVCTop10.length)
            return '';

        var baseMinima = "";

        this.graficoBVCTop10.forEach(element => {


            if (element.BaseMinimaShareDesejo != '') {
                baseMinima = element.BaseMinimaShareDesejo
            }

            if (element.BaseMinimaEfeitos != '') {
                baseMinima = element.BaseMinimaEfeitos
            }

            if (element.BaseMinimaEquity != '') {
                baseMinima = element.BaseMinimaEquity
            }
        });

        return baseMinima;
    }


}