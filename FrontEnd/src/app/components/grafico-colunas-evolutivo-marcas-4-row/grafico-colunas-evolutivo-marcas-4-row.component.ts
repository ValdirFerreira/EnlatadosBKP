import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';

import { GraficoColunaModel } from 'src/app/models/grafico-coluna/grafico-coluna';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';


@Component({
    selector: 'grafico-colunas-evolutivo-marcas-4-row',
    templateUrl: './grafico-colunas-evolutivo-marcas-4-row.component.html',
    styleUrls: ['./grafico-colunas-evolutivo-marcas-4-row.component.scss']
})
export class GraficoColunasEvolutivoMarcas4RowComponent implements OnInit {

    @Input() graficoColunaModel1: GraficoColunaModel;
    @Input() graficoColunaModel2: GraficoColunaModel;
    @Input() graficoColunaModel3: GraficoColunaModel;
    @Input() graficoColunaModel4: GraficoColunaModel;
    @Input() graficoColunaModel5: GraficoColunaModel;

    @Output('ondaColuna1') ondaColuna1: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna2') ondaColuna2: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna3') ondaColuna3: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna4') ondaColuna4: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna5') ondaColuna5: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

    @Output('marcaGrafico') marcaGrafico: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

    onda1: PadraoComboFiltro;
    onda2: PadraoComboFiltro;
    onda3: PadraoComboFiltro;
    onda4: PadraoComboFiltro;

    descInicial: string = "Ninho";

    public ModelMarcas: PadraoComboFiltro;

    constructor(
        public filtroService: FiltroGlobalService,
    ) { }

    ngOnInit(): void {
        this.FiltroOndas();
    }


    onchangeMarcaGrafico(item: PadraoComboFiltro) {
        this.marcaGrafico.emit(item);
    }

    onchangeOndaColuna1(item: PadraoComboFiltro) {
        this.ondaColuna1.emit(item);
    }
    onchangeOndaColuna2(item: PadraoComboFiltro) {
        this.ondaColuna2.emit(item);
    }
    onchangeOndaColuna3(item: PadraoComboFiltro) {
        this.ondaColuna3.emit(item);
    }
    onchangeOndaColuna4(item: PadraoComboFiltro) {
        this.ondaColuna4.emit(item);
    }
    onchangeOndaColuna5(item: PadraoComboFiltro) {
        this.ondaColuna5.emit(item);
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

        if (graficoColunaModel.TesteSigOutros) {

            if (graficoColunaModel.TesteSigOutros == 'MAIOR')
                return 'sig-positive';

            if (graficoColunaModel.TesteSigOutros == 'MENOR')
                return 'sig-negative';

            if (graficoColunaModel.TesteSigOutros == 'IGUAL')
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



    montaImageMarca() {

        if (this.ModelMarcas != undefined) {
            return "assets/marcas/" + this.ModelMarcas.IdItem + ".svg";
        }
        else {
            if (this.filtroService.ModelDenominators && this.filtroService.ModelDenominators.IdItem > 1) {
                this.descInicial = this.filtroService.listaMarcas[0].DescItem;
                return "assets/marcas/"+this.filtroService.listaMarcas[0].IdItem+".svg";
            }
            else
                return "assets/marcas/13001.svg";
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