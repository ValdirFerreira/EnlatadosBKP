import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';
import { ComunicacaoLikeDislikeModel } from 'src/app/models/grafico-coluna/ComunicacaoLikeDislikeModel';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';

@Component({
    selector: 'comunicacao-like-dislike',
    templateUrl: './comunicacao-like-dislike.component.html',
    styleUrls: ['./comunicacao-like-dislike.component.scss']
})
export class ComunicacaoLikeDislikeComponent implements OnInit {

    @Input() comunicacaoLikeDislikeModel1: ComunicacaoLikeDislikeModel;
    @Input() comunicacaoLikeDislikeModel2: ComunicacaoLikeDislikeModel;
    @Input() comunicacaoLikeDislikeModel3: ComunicacaoLikeDislikeModel;
    @Input() comunicacaoLikeDislikeModel4: ComunicacaoLikeDislikeModel;
    @Input() comunicacaoLikeDislikeModel5: ComunicacaoLikeDislikeModel;

    @Output('marcaColuna1') marcaColuna1: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna2') marcaColuna2: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna3') marcaColuna3: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna4') marcaColuna4: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna5') marcaColuna5: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

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

    ajusteColuna(valor: any): string {

        if (valor == null || valor == undefined || valor <= 0) {
            return '8%';
        }

        if (valor < 8) {
            return '8%';
        }

        if (valor > 100) {
            return '100%';
        }

        return valor + '%';
    }

    validaSig(sig: string): string {

        if (!sig) {
            return 'sig-vazio';
        }

        sig = sig.toUpperCase();

        if (sig == 'MAIOR') {
            return 'sig-positive';
        }

        if (sig == 'MENOR') {
            return 'sig-negative';
        }

        if (sig == 'IGUAL') {
            return 'sig-vazio';
        }

        return 'sig-vazio';
    }

    validaColunaSemDados(model: ComunicacaoLikeDislikeModel): boolean {

        if (!model) {
            return false;
        }

        if (
            model.PercGostei == 0 &&
            model.PercGosteiPouco == 0 &&
            model.PercNenhum == 0 &&
            model.PercNaoGostei == 0 &&
            model.PercNaoGosteiPouco == 0
        ) {
            return false;
        }

        return true;
    }

    verificaMsgBase(): string {

        if (this.comunicacaoLikeDislikeModel1?.BaseMinima) {
            return this.comunicacaoLikeDislikeModel1.BaseMinima;
        }

        if (this.comunicacaoLikeDislikeModel2?.BaseMinima) {
            return this.comunicacaoLikeDislikeModel2.BaseMinima;
        }

        if (this.comunicacaoLikeDislikeModel3?.BaseMinima) {
            return this.comunicacaoLikeDislikeModel3.BaseMinima;
        }

        if (this.comunicacaoLikeDislikeModel4?.BaseMinima) {
            return this.comunicacaoLikeDislikeModel4.BaseMinima;
        }

        if (this.comunicacaoLikeDislikeModel5?.BaseMinima) {
            return this.comunicacaoLikeDislikeModel5.BaseMinima;
        }

        return '';
    }
}