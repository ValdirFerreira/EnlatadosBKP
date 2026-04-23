import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';
import { GraficoColunaModel } from 'src/app/models/grafico-coluna/grafico-coluna';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';


@Component({
    selector: 'grafico-colunas-marcas-4-row',
    templateUrl: './grafico-colunas-marcas-4-row.component.html',
    styleUrls: ['./grafico-colunas-marcas-4-row.component.scss']
})
export class GraficoColunasMarcas4RowComponent implements OnInit {

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

   
    ajusteDeColuna0(PercPrompeted: any, PercOM: any, PercTOM: any, PercOutros: any, t:any): string {
        const PercTotal = PercPrompeted + PercOM + PercTOM + PercOutros;
        let retorno;
    
        if (PercPrompeted > 4) {
            retorno = (PercPrompeted / PercTotal) * 100;
        } else {
            return '4%';
        }
    
        return retorno.toFixed(2) + '%';
    }
    
    ajusteDeColuna1(PercPrompeted: any, PercOM: any, PercTOM: any, PercOutros: any,t:any): string {
        const PercTotal = PercPrompeted + PercOM + PercTOM + PercOutros;
        let retorno;
    
        if (PercPrompeted > 4) {
            retorno = (PercPrompeted / PercTotal) * 100;
        } else {
            return '4%';
        }
    
        return retorno.toFixed(2) + '%';
    }
    
    ajusteDeColuna2(PercPrompeted: any, PercOM: any, PercTOM: any, PercOutros: any,t:any): string {
        const PercTotal = PercPrompeted + PercOM + PercTOM + PercOutros;
        let retorno;
    
        if (PercOM > 4) {
            retorno = (PercOM / PercTotal) * 100;
        } else {
            return '4%';
        }
    
        return retorno.toFixed(2) + '%';
    }
    
    ajusteDeColuna3(PercPrompeted: any, PercOM: any, PercTOM: any, PercOutros: any,t:any): string {
        const PercTotal = PercPrompeted + PercOM + PercTOM + PercOutros;
        let retorno;
    
        if (PercTOM > 4) {
            retorno = (PercTOM / PercTotal) * 100;
        } else {
            return '4%';
        }
    
        return retorno.toFixed(2) + '%';
    }
    
    ajusteDeColuna4(PercPrompeted: any, PercOM: any, PercTOM: any, PercOutros: any,t:any): string {
        const PercTotal = PercPrompeted + PercOM + PercTOM + PercOutros;
        let retorno;
    
        if (PercOutros > 4) {
            retorno = (PercOutros / PercTotal) * 100;
        } else {
            return '4%';
        }
    
        return retorno.toFixed(2) + '%';
    }

    validaSigTom(graficoColunaModel: GraficoColunaModel) {

        if (graficoColunaModel.TesteSIGTOM) {

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

        if (graficoColunaModel.TesteSIGOM) {
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

        if (graficoColunaModel.TesteSIGPrompeted) {
            if (graficoColunaModel.TesteSIGPrompeted == 'MAIOR')
                return 'sig-positive';

            if (graficoColunaModel.TesteSIGPrompeted == 'MENOR')
                return 'sig-negative';

            if (graficoColunaModel.TesteSIGPrompeted == 'IGUAL')
                return 'sig-vazio';
        }
        return 'sig-vazio';
    }

    validaSigPercTotal(graficoColunaModel: GraficoColunaModel) {

        if (graficoColunaModel.TesteSigTotal) {
            if (graficoColunaModel.TesteSigTotal == 'MAIOR')
                return 'sig-positive';

            if (graficoColunaModel.TesteSigTotal == 'MENOR')
                return 'sig-negative';

            if (graficoColunaModel.TesteSigTotal == 'IGUAL')
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

    validaSigOutros(graficoColunaModel: GraficoColunaModel) {

        if (graficoColunaModel.PercTotal != 0) {

            if (graficoColunaModel.TesteSigOutros == 'MAIOR')
                return 'sig-positive';

            if (graficoColunaModel.TesteSigOutros == 'MENOR')
                return 'sig-negative';

            if (graficoColunaModel.TesteSigOutros == 'IGUAL')
                return 'sig-vazio';
        }
        return 'sig-vazio';
    }

    validaColunaSemDados(graficoColunaModel: GraficoColunaModel) {
        if (graficoColunaModel.PercPrompeted == 0 && graficoColunaModel.PercOM == 0 && graficoColunaModel.PercTOM == 0
            && graficoColunaModel.PercOutros == 0   && graficoColunaModel.PercTotal == 0  
        ) {
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