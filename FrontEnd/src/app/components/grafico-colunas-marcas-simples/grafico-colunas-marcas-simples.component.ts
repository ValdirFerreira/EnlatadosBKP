import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';
import { GraficoColunaModel } from 'src/app/models/grafico-coluna/grafico-coluna';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';


@Component({
    selector: 'grafico-colunas-marcas-simples',
    templateUrl: './grafico-colunas-marcas-simples.component.html',
    styleUrls: ['./grafico-colunas-marcas-simples.component.scss']
})
export class GraficoColunasMarcasSimplesComponent implements OnInit {

    @Input() graficoColunaModel1: GraficoColunaModel;
    @Input() graficoColunaModel2: GraficoColunaModel;
    @Input() graficoColunaModel3: GraficoColunaModel;
    @Input() graficoColunaModel4: GraficoColunaModel;
    @Input() graficoColunaModel5: GraficoColunaModel;

    @Output('marcaColuna1') marcaColuna1: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna2') marcaColuna2: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna3') marcaColuna3: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna4') marcaColuna4: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna5') marcaColuna5: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

    public ModelMarcasColuna1: PadraoComboFiltro;

    ngOnInit(): void {

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

    ajusteDeColuna1(PercPrompeted: any, PercOM: any, PercTOM: any, PercTotal: any) {
        var retorno;
        if (PercOM >= 15) {
            retorno = PercPrompeted;
        } else retorno = '85%'

        if (PercTOM >= 10) {
        } else retorno = '90%'

        if (PercOM >= 15 && PercTOM >= 10) {
            retorno = PercPrompeted
        } else retorno = '75'

        var total = PercOM + PercTOM;
        var total3colunas = PercOM + PercTOM + PercPrompeted;

        if (total3colunas < 100) {
            retorno = 100 - total
        } else {
            retorno = PercPrompeted;
        }

        if (PercOM == 0 && PercTOM == 0) {
            retorno = '100'
        }

        if (PercOM > 0 && PercTOM == 0 || PercOM == 0 && PercTOM > 0) {
            retorno = '50';
        }

        if (PercTotal == 0) {
            if (PercOM > 0 && PercTOM > 0 || PercOM > 0 && PercTOM > 0) {
                retorno = '50';
            }
        }
        return retorno + '%';
    }

    ajusteDeColuna2(PercPrompeted: any, PercOM: any, PercTOM: any, PercTotal: any): string {
        var retorno;

        if (PercOM >= 15) {
            retorno = PercOM;
        } else {
            retorno = '15';
        }

        if (PercPrompeted == 0 && PercTOM == 0) {
            retorno = '100';
        }

        if (PercPrompeted > 0 && PercTOM == 0 || PercPrompeted == 0 && PercTOM > 0) {
            retorno = '50';
        }

        if (PercTotal == 0) {
            if (PercPrompeted > 0 && PercTOM > 0 || PercPrompeted > 0 && PercTOM > 0) {
                retorno = '25';
            }
        }

        return retorno + '%';
    }

    ajusteDeColuna3(PercPrompeted: any, PercOM: any, PercTOM: any, PercTotal: any): string {
        var retorno;

        if (PercTOM >= 10) {
            retorno = PercTOM;
        } else {
            retorno = '10';
        }

        if (PercPrompeted == 0 && PercOM == 0) {
            retorno = '100';
        }

        if (PercPrompeted > 0 && PercOM == 0 || PercPrompeted == 0 && PercOM > 0) {
            retorno = '50';
        }

        if (PercTotal == 0) {
            if (PercPrompeted > 0 && PercOM > 0 || PercPrompeted > 0 && PercOM > 0) {
                retorno = '25';
            }
        }

        return retorno + '%';
    }

    validaSigTom(graficoColunaModel: GraficoColunaModel) {

        if (graficoColunaModel.PercTotal != 0) {

            if (graficoColunaModel.TesteSIGTOM == 'MAIOR')
                return 'sig-positive';

            if (graficoColunaModel.TesteSIGTOM == 'MENOR')
                return 'sig-negative';

            if (graficoColunaModel.TesteSIGTOM == 'IGUAL')
                return 'sig-vazio';
        }
        return 'sig-vazio';
    }

    validaSigOm(graficoColunaModel: GraficoColunaModel) {

        if (graficoColunaModel.PercTotal != 0) {
            if (graficoColunaModel.TesteSIGOM == 'MAIOR')
                return 'sig-positive';

            if (graficoColunaModel.TesteSIGOM == 'MENOR')
                return 'sig-negative';

            if (graficoColunaModel.TesteSIGOM == 'IGUAL')
                return 'sig-vazio';
        }
        return 'sig-vazio';
    }

    validaSigPrompeted(graficoColunaModel: GraficoColunaModel) {

        if (graficoColunaModel.PercTotal != 0) {
            if (graficoColunaModel.TesteSIGPrompeted == 'MAIOR')
                return 'sig-positive';

            if (graficoColunaModel.TesteSIGPrompeted == 'MENOR')
                return 'sig-negative';

            if (graficoColunaModel.TesteSIGPrompeted == 'IGUAL')
                return 'sig-vazio';
        }
        return 'sig-vazio';
    }


    validaSigTotal(graficoColunaModel: GraficoColunaModel) {

        // if (graficoColunaModel.PercTotal != 0) {
            if (graficoColunaModel.TesteSigTotal == 'MAIOR')
                return 'sig-positive';

            if (graficoColunaModel.TesteSigTotal == 'MENOR')
                return 'sig-negative';

            if (graficoColunaModel.TesteSigTotal == 'IGUAL')
                return 'sig-vazio';
        // }
        return 'sig-vazio';
    }

    validaColunaSemDados(graficoColunaModel: GraficoColunaModel) {
        if (graficoColunaModel.PercPrompeted == 0 && graficoColunaModel.PercOM == 0 && graficoColunaModel.PercTOM == 0) {
            return false;
        }
        else {
            return true;
        }
    }


    verificaMsgBase() {
        if (this.graficoColunaModel1?.BaseMinima) {
            return this.graficoColunaModel1.BaseMinima;
        }

        if (this.graficoColunaModel2?.BaseMinima) {
            return this.graficoColunaModel2.BaseMinima;
        }

        if (this.graficoColunaModel3?.BaseMinima) {
            return this.graficoColunaModel3.BaseMinima;
        }

        if (this.graficoColunaModel4?.BaseMinima) {
            return this.graficoColunaModel4.BaseMinima;
        }

        if (this.graficoColunaModel5?.BaseMinima) {
            return this.graficoColunaModel5.BaseMinima;
        }

        return '';
    }

}