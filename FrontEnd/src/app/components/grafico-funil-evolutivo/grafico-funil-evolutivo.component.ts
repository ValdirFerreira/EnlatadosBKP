import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';
import { GraficoColunaModel } from 'src/app/models/grafico-coluna/grafico-coluna';
import { GraficoFunil } from 'src/app/models/grafico-funil/GraficoFunil';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';


@Component({
    selector: 'grafico-funil-evolutivo',
    templateUrl: './grafico-funil-evolutivo.component.html',
    styleUrls: ['./grafico-funil-evolutivo.component.scss']
})
export class GraficoFunilEvolutivoComponent implements OnInit {

    @Input() graficoFunil1: GraficoFunil;
    @Input() graficoFunil2: GraficoFunil;
    @Input() graficoFunil3: GraficoFunil;
    @Input() graficoFunil4: GraficoFunil;

    @Output('ondaColuna1') ondaColuna1: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna2') ondaColuna2: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna3') ondaColuna3: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna4') ondaColuna4: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

    @Output('marcaGrafico') marcaGrafico: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

    onda1: PadraoComboFiltro;
    onda2: PadraoComboFiltro;
    onda3: PadraoComboFiltro;
    onda4: PadraoComboFiltro;

    descInicial: string = "Nestlé";

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

    montaImageMarca() {
        if (this.ModelMarcas != undefined) {
            this.descInicial = this.ModelMarcas.DescItem;

            return "assets/marcas/" + this.ModelMarcas.IdItem + ".svg";
        }
        else {
            // return "assets/marcas/1.svg";

            this.descInicial = this.filtroService.listaMarcas[0].DescItem;
            return "assets/marcas/" + this.filtroService.listaMarcas[0].IdItem + ".svg";
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


    validaSigConhecimento(graficoFunil1: GraficoFunil) {

        if (graficoFunil1.BaseAbsAtual != 0) {
            if (graficoFunil1.ConhecimentoTesteAtual == 'MAIOR')
                return 'sig-positive';

            if (graficoFunil1.ConhecimentoTesteAtual == 'MENOR')
                return 'sig-negative';

            if (graficoFunil1.ConhecimentoTesteAtual == 'IGUAL')
                return 'sig-vazio';
        }
        return 'sig-vazio';
    }

    validaSigConsideracao(graficoFunil1: GraficoFunil) {

        if (graficoFunil1.BaseAbsAtual != 0) {
            if (graficoFunil1.ConsideracaoTesteAtual == 'MAIOR')
                return 'sig-positive';

            if (graficoFunil1.ConsideracaoTesteAtual == 'MENOR')
                return 'sig-negative';

            if (graficoFunil1.ConsideracaoTesteAtual == 'IGUAL')
                return 'sig-vazio';
        }
        return 'sig-vazio';
    }

    validaSigUso(graficoFunil1: GraficoFunil) {

        if (graficoFunil1.BaseAbsAtual != 0) {
            if (graficoFunil1.UsoTesteAtual == 'MAIOR')
                return 'sig-positive';

            if (graficoFunil1.UsoTesteAtual == 'MENOR')
                return 'sig-negative';

            if (graficoFunil1.UsoTesteAtual == 'IGUAL')
                return 'sig-vazio';
        }
        return 'sig-vazio';
    }

    validaSigPreferencia(graficoFunil1: GraficoFunil) {

        if (graficoFunil1.BaseAbsAtual != 0) {
            if (graficoFunil1.PreferenciaTesteAtual == 'MAIOR')
                return 'sig-positive';

            if (graficoFunil1.PreferenciaTesteAtual == 'MENOR')
                return 'sig-negative';

            if (graficoFunil1.PreferenciaTesteAtual == 'IGUAL')
                return 'sig-vazio';
        }
        return 'sig-vazio';
    }

    validaSigLoyalty(graficoFunil1: GraficoFunil) {

        if (graficoFunil1.BaseAbsAtual != 0) {
            if (graficoFunil1.LoyaltyTesteAtual == 'MAIOR')
                return 'sig-positive';

            if (graficoFunil1.LoyaltyTesteAtual == 'MENOR')
                return 'sig-negative';

            if (graficoFunil1.LoyaltyTesteAtual == 'IGUAL')
                return 'sig-vazio';
        }
        return 'sig-vazio';
    }

    validaSigCalc6(graficoFunil1: GraficoFunil) {

        if (graficoFunil1.BaseAbsAtual != 0) {
            if (graficoFunil1.Calc6TesteAtual == 'MAIOR')
                return 'sig-positive';

            if (graficoFunil1.Calc6TesteAtual == 'MENOR')
                return 'sig-negative';

            if (graficoFunil1.Calc6TesteAtual == 'IGUAL')
                return 'sig-vazio';
        }
        return 'sig-vazio';
    }

    validaSigCalc7(graficoFunil1: GraficoFunil) {

        if (graficoFunil1.BaseAbsAtual != 0) {
            if (graficoFunil1.Calc7TesteAtual == 'MAIOR')
                return 'sig-positive';

            if (graficoFunil1.Calc7TesteAtual == 'MENOR')
                return 'sig-negative';

            if (graficoFunil1.Calc7TesteAtual == 'IGUAL')
                return 'sig-vazio';
        }
        return 'sig-vazio';
    }

    setCorGrafico() {
        if (!this.ModelMarcas) {

            if (this.filtroService.listaMarcas)
                return this.filtroService.listaMarcas[0].CorItem;
            else
                return '#ffffff';
        }
        else
            return this.ModelMarcas.CorItem;
    }

    definePorcentagem(valor) {
        return valor + "%";
    }

    verificaMsgBase() {
        if (this.graficoFunil1?.BaseMinimaAtual) {
            return this.graficoFunil1.BaseMinimaAtual;
        }
        if (this.graficoFunil1?.BaseMinimaAnterior) {
            return this.graficoFunil1.BaseMinimaAnterior;
        }

        if (this.graficoFunil2?.BaseMinimaAtual) {
            return this.graficoFunil2.BaseMinimaAtual;
        }
        if (this.graficoFunil2?.BaseMinimaAnterior) {
            return this.graficoFunil2.BaseMinimaAnterior;
        }


        if (this.graficoFunil3?.BaseMinimaAtual) {
            return this.graficoFunil3.BaseMinimaAtual;
        }
        if (this.graficoFunil3?.BaseMinimaAnterior) {
            return this.graficoFunil3.BaseMinimaAnterior;
        }

        if (this.graficoFunil4?.BaseMinimaAtual) {
            return this.graficoFunil4.BaseMinimaAtual;
        }
        if (this.graficoFunil4?.BaseMinimaAnterior) {
            return this.graficoFunil4.BaseMinimaAnterior;
        }



        return '';
    }

}