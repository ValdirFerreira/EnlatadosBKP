import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
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

    @Output() marcaColuna1 = new EventEmitter<PadraoComboFiltro>();
    @Output() marcaColuna2 = new EventEmitter<PadraoComboFiltro>();
    @Output() marcaColuna3 = new EventEmitter<PadraoComboFiltro>();
    @Output() marcaColuna4 = new EventEmitter<PadraoComboFiltro>();
    @Output() marcaColuna5 = new EventEmitter<PadraoComboFiltro>();
    @Output() marcaColuna6 = new EventEmitter<PadraoComboFiltro>();
    @Output() marcaColuna7 = new EventEmitter<PadraoComboFiltro>();
    @Output() marcaColuna8 = new EventEmitter<PadraoComboFiltro>();

    constructor(
        private translate: TranslateService,
        public filtroService: FiltroGlobalService,
    ) { }

    ngOnInit(): void { }

    onchangeMarcaColuna1(item: PadraoComboFiltro) { this.marcaColuna1.emit(item); }
    onchangeMarcaColuna2(item: PadraoComboFiltro) { this.marcaColuna2.emit(item); }
    onchangeMarcaColuna3(item: PadraoComboFiltro) { this.marcaColuna3.emit(item); }
    onchangeMarcaColuna4(item: PadraoComboFiltro) { this.marcaColuna4.emit(item); }
    onchangeMarcaColuna5(item: PadraoComboFiltro) { this.marcaColuna5.emit(item); }
    onchangeMarcaColuna6(item: PadraoComboFiltro) { this.marcaColuna6.emit(item); }
    onchangeMarcaColuna7(item: PadraoComboFiltro) { this.marcaColuna7.emit(item); }
    onchangeMarcaColuna8(item: PadraoComboFiltro) { this.marcaColuna8.emit(item); }

    // 🔹 Função genérica (evita repetição)
    private validaSig(valor: string, base: number): string {
        if (base !== 0) {
            if (valor === 'MAIOR') return 'sig-positive';
            if (valor === 'MENOR') return 'sig-negative';
            if (valor === 'IGUAL') return 'sig-vazio';
        }
        return 'sig-vazio';
    }

    validaSigConhecimento(g: GraficoFunil) {
        return this.validaSig(g.ConhecimentoTesteAtual, g.BaseAbsAtual);
    }

    validaSigConsideracao(g: GraficoFunil) {
        return this.validaSig(g.ConsideracaoTesteAtual, g.BaseAbsAtual);
    }

    validaSigUso(g: GraficoFunil) {
        return this.validaSig(g.UsoTesteAtual, g.BaseAbsAtual);
    }

    validaSigPreferencia(g: GraficoFunil) {
        return this.validaSig(g.PreferenciaTesteAtual, g.BaseAbsAtual);
    }

    validaSigLoyalty(g: GraficoFunil) {
        return this.validaSig(g.LoyaltyTesteAtual, g.BaseAbsAtual);
    }

    // 🔥 NOVOS (Calc6 e Calc7)
    validaSigCalc6(g: GraficoFunil) {
        return this.validaSig(g.Calc6TesteAtual, g.BaseAbsAtual);
    }

    validaSigCalc7(g: GraficoFunil) {
        return this.validaSig(g.Calc7TesteAtual, g.BaseAbsAtual);
    }

    setCorGrafico(g: GraficoFunil) {
        return g?.CorSiteMarca;
    }

    definePorcentagem(valor: number): string {
        return `${valor}%`;
    }

    verificaMsgBase(): string {
        const lista = [
            this.graficoFunil1, this.graficoFunil2, this.graficoFunil3, this.graficoFunil4,
            this.graficoFunil5, this.graficoFunil6, this.graficoFunil7, this.graficoFunil8
        ];

        for (let g of lista) {
            if (g?.BaseMinimaAtual) return g.BaseMinimaAtual;
            if (g?.BaseMinimaAnterior) return g.BaseMinimaAnterior;
        }

        return '';
    }

    validaBase(g: GraficoFunil): string {
        if (this.filtroService.ModelOnda && this.filtroService.ModelOnda.IdItem == 1) {
            return g.BaseAbsAnterior <= 0 ? "N/A" : g.PeriodoAnt;
        }
        return g.PeriodoAnt;
    }
}