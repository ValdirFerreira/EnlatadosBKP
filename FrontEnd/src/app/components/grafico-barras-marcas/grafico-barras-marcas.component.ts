import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';
import { GraficoColunaModel } from 'src/app/models/grafico-coluna/grafico-coluna';
import { GraficoImagemPosicionamento } from 'src/app/models/grafico-Imagem-posicionamento/GraficoImagemPosicionamento';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';


@Component({
    selector: 'grafico-barras-marcas',
    templateUrl: './grafico-barras-marcas.component.html',
    styleUrls: ['./grafico-barras-marcas.component.scss']
})
export class GraficoBarrasMarcasComponent implements OnInit {



    @Input() graficoColunaModel1: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel2: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel3: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel4: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel5: Array<GraficoImagemPosicionamento>;

    @Output('marcaColuna1') marcaColuna1: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna2') marcaColuna2: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna3') marcaColuna3: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna4') marcaColuna4: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna5') marcaColuna5: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

    qtdLinhas: number = 0;

    ngOnInit(): void {

        this.qtdLinhas = this.graficoColunaModel1.length - 1;
    }

    ngOnChanges(changes: any) {

        this.qtdLinhas = this.graficoColunaModel1.length - 1;

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



    montaCorBarra(item: GraficoImagemPosicionamento) {
        return 'border:1px solid ' + item.CorSiteMarca + ' !important;';

    }


    montaCorBarraTesteSig(item: GraficoImagemPosicionamento) {
        if (item.PercAtual <= -20 || item.PercAtual > 20)
            return 'border:1px solid ' + item.CorSiteMarca + ' !important;';
        else
            return '';

    }


    verificaMsgBase() {
        // if (!this.graficoColunaModel1.length)
        //     return '';

        if (this.verificaObjeto(this.graficoColunaModel1)) {
            return this.graficoColunaModel1[0].BaseMinima;
        }

        if (this.verificaObjeto(this.graficoColunaModel2)) {
            return this.graficoColunaModel2[0].BaseMinima;
        }

        if (this.verificaObjeto(this.graficoColunaModel3)) {
            return this.graficoColunaModel3[0].BaseMinima;
        }

        if (this.verificaObjeto(this.graficoColunaModel4)) {
            return this.graficoColunaModel4[0].BaseMinima;
        }

        if (this.verificaObjeto(this.graficoColunaModel5)) {
            return this.graficoColunaModel5[0].BaseMinima;
        }

        return '';
    }


    verificaObjeto(obj: Array<GraficoImagemPosicionamento>) {
        if (obj != undefined && obj != null && obj.length > 0) {
            var retorno = false;
            obj.forEach(item => {
                if (item.BaseMinima != "") {
                    retorno = true;
                }
            });
            return retorno;
        }
        else {
            return false;
        }
    }

}