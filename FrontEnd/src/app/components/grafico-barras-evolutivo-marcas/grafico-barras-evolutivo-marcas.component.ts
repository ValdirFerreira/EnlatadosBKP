import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';
import { GraficoColunaModel } from 'src/app/models/grafico-coluna/grafico-coluna';
import { GraficoImagemPosicionamento } from 'src/app/models/grafico-Imagem-posicionamento/GraficoImagemPosicionamento';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';


@Component({
    selector: 'grafico-barras-evolutivo-marcas',
    templateUrl: './grafico-barras-evolutivo-marcas.component.html',
    styleUrls: ['./grafico-barras-evolutivo-marcas.component.scss']
})
export class GraficoBarrasEvolutivoMarcasComponent implements OnInit {


    @Input() graficoColunaModel1: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel2: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel3: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel4: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel5: Array<GraficoImagemPosicionamento>;


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

    descInicial: string = "Nestlé";

    public ModelMarcas: PadraoComboFiltro;

    constructor(
        public filtroService: FiltroGlobalService,
    ) { }

    qtdLinhas: number = 0;

    ngOnInit(): void {
        this.qtdLinhas = 24;
        this.FiltroOndas();
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
        if (!this.graficoColunaModel1.length)
            return '';

        if (this.validaObjeto(this.graficoColunaModel1) && this.graficoColunaModel1?.[0].BaseMinima) {
            return this.graficoColunaModel1[0].BaseMinima;
        }

        if (this.validaObjeto(this.graficoColunaModel2) &&this.graficoColunaModel2?.[0].BaseMinima) {
            return this.graficoColunaModel2[0].BaseMinima;
        }

        if (this.validaObjeto(this.graficoColunaModel3) &&this.graficoColunaModel3?.[0].BaseMinima) {
            return this.graficoColunaModel3[0].BaseMinima;
        }

        if ( this.validaObjeto(this.graficoColunaModel4) &&this.graficoColunaModel4?.[0].BaseMinima) {
            return this.graficoColunaModel4[0].BaseMinima;
        }

        if (this.validaObjeto(this.graficoColunaModel5) &&this.graficoColunaModel5?.[0].BaseMinima) {
            return this.graficoColunaModel5[0].BaseMinima;
        }

        return '';
    }


    validaObjeto(obj: Array<GraficoImagemPosicionamento>) {
        if (obj != null && obj != undefined && obj.length > 0) {
            return true;
        }

        return false;
    }

}