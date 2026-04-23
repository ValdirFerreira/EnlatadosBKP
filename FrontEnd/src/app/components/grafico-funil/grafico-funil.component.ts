import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { GraficoColunaModel } from 'src/app/models/grafico-coluna/grafico-coluna';
import { GraficoFunil } from 'src/app/models/grafico-funil/GraficoFunil';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';


@Component({
    selector: 'grafico-funil',
    templateUrl: './grafico-funil.component.html',
    styleUrls: ['./grafico-funil.component.scss']
})
export class GraficoFunilComponent implements OnInit {

    @Input() graficoFunil1: GraficoFunil;
    @Input() graficoFunil2: GraficoFunil;
    @Input() graficoFunil3: GraficoFunil;
    @Input() graficoFunil4: GraficoFunil;
    @Input() graficoFunil5: GraficoFunil;
    @Input() graficoFunil6: GraficoFunil;
    @Input() graficoFunil7: GraficoFunil;
    @Input() graficoFunil8: GraficoFunil;

    @Output('marcaColuna1') marcaColuna1: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna2') marcaColuna2: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna3') marcaColuna3: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna4') marcaColuna4: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna5') marcaColuna5: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna6') marcaColuna6: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna7') marcaColuna7: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna8') marcaColuna8: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();


    constructor(
        private translate: TranslateService, public filtroService: FiltroGlobalService,
    ) { }

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

    onchangeMarcaColuna6(item: PadraoComboFiltro) {
        this.marcaColuna6.emit(item);
    }

    onchangeMarcaColuna7(item: PadraoComboFiltro) {
        this.marcaColuna7.emit(item);
    }

    onchangeMarcaColuna8(item: PadraoComboFiltro) {
        this.marcaColuna8.emit(item);
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


    setCorGrafico(graficoFunil: GraficoFunil) {
        return graficoFunil.CorSiteMarca;
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

        if (this.graficoFunil5?.BaseMinimaAtual) {
            return this.graficoFunil5.BaseMinimaAtual;
        }
        if (this.graficoFunil5?.BaseMinimaAnterior) {
            return this.graficoFunil5.BaseMinimaAnterior;
        }

        if (this.graficoFunil6?.BaseMinimaAtual) {
            return this.graficoFunil6.BaseMinimaAtual;
        }
        if (this.graficoFunil6?.BaseMinimaAnterior) {
            return this.graficoFunil6.BaseMinimaAnterior;
        }

        if (this.graficoFunil7?.BaseMinimaAtual) {
            return this.graficoFunil7.BaseMinimaAtual;
        }
        if (this.graficoFunil7?.BaseMinimaAnterior) {
            return this.graficoFunil7.BaseMinimaAnterior;
        }

        if (this.graficoFunil8?.BaseMinimaAtual) {
            return this.graficoFunil8.BaseMinimaAtual;
        }
        if (this.graficoFunil8?.BaseMinimaAnterior) {
            return this.graficoFunil8.BaseMinimaAnterior;
        }


        return '';
    }


    validaBase(graficoFunil: GraficoFunil) {

        if (this.filtroService.ModelOnda && this.filtroService.ModelOnda.IdItem == 1) // SE 2019 custom para Nestlé Formula Infantil
        {
            if (graficoFunil.BaseAbsAnterior <= 0) {
                return "N/A";
            }
            else {
                return graficoFunil.PeriodoAnt;
            }
        }
        else {
            return graficoFunil.PeriodoAnt;
        }

    }

}