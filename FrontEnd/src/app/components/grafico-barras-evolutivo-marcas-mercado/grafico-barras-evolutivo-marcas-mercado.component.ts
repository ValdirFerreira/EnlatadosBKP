import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';
import { GraficoColunaModel } from 'src/app/models/grafico-coluna/grafico-coluna';
import { GraficoImagemPosicionamento } from 'src/app/models/grafico-Imagem-posicionamento/GraficoImagemPosicionamento';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';


@Component({
    selector: 'grafico-barras-evolutivo-marcas-mercado',
    templateUrl: './grafico-barras-evolutivo-marcas-mercado.component.html',
    styleUrls: ['./grafico-barras-evolutivo-marcas-mercado.component.scss']
})


export class GraficoBarrasEvolutivoMarcasMercadoComponent implements OnInit {


    @Input() graficoColunaModel1: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel2: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel3: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel4: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel5: Array<GraficoImagemPosicionamento>;


    @Output('ondaColuna10') ondaColuna10: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna20') ondaColuna20: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna30') ondaColuna30: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna40') ondaColuna40: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna50') ondaColuna50: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

    @Output('marcaGrafico0') marcaGrafico: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

    graficoColunaDescricao: Array<GraficoImagemPosicionamento>;

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
        // this.qtdLinhas = 19;
        this.FiltroOndas();
    }

    ngOnChanges(changes: any) {
        this.verificaDescricao();

    }


    porcBarraPositiva(valor) {

        if (valor >= 0.1 && valor < 1)
        return "1%";

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
        this.ondaColuna10.emit(item);
    }
    onchangeOndaColuna2(item: PadraoComboFiltro) {
        this.ondaColuna20.emit(item);
    }
    onchangeOndaColuna3(item: PadraoComboFiltro) {
        this.ondaColuna30.emit(item);
    }
    onchangeOndaColuna4(item: PadraoComboFiltro) {
        this.ondaColuna40.emit(item);
    }
    onchangeOndaColuna5(item: PadraoComboFiltro) {
        this.ondaColuna50.emit(item);
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
        // if (!this.graficoColunaModel1 || !this.graficoColunaModel1.length)
        //     return '';

        if (this.validaObjeto(this.graficoColunaModel1)) {
            return this.graficoColunaModel1[0].BaseMinima;
        }

        if (this.validaObjeto(this.graficoColunaModel2)) {
            return this.graficoColunaModel2[0].BaseMinima;
        }


        if (this.validaObjeto(this.graficoColunaModel3)) {
            return this.graficoColunaModel3[0].BaseMinima;
        }

        if (this.validaObjeto(this.graficoColunaModel4)) {
            return this.graficoColunaModel4[0].BaseMinima;
        }

        if (this.validaObjeto(this.graficoColunaModel5)) {
            return this.graficoColunaModel5[0].BaseMinima;
        }

        return '';
    }


    verificaMsgBaseMinima(grafico: Array<GraficoImagemPosicionamento>) {

        var descricao = false;

        if (grafico) {
            grafico.forEach(item => {
                if (item.BaseMinima != '') {
                    descricao = true;
                }
            }
            );
        }

        return descricao;
    }




    verificaDescricao() {

        if (this.validaObjeto(this.graficoColunaModel1)) {
            this.graficoColunaDescricao = this.graficoColunaModel1;
            this.qtdLinhas = this.graficoColunaModel1.length - 1;
        }

        if (this.validaObjeto(this.graficoColunaModel2)) {
            this.graficoColunaDescricao = this.graficoColunaModel2;
            this.qtdLinhas = this.graficoColunaModel2.length - 1;
        }

        if (this.validaObjeto(this.graficoColunaModel3)) {
            this.graficoColunaDescricao = this.graficoColunaModel3;
            this.qtdLinhas = this.graficoColunaModel3.length - 1;
        }

        if (this.validaObjeto(this.graficoColunaModel4)) {
            this.graficoColunaDescricao = this.graficoColunaModel4;
            this.qtdLinhas = this.graficoColunaModel4.length - 1;
        }

        if (this.validaObjeto(this.graficoColunaModel5)) {
            this.graficoColunaDescricao = this.graficoColunaModel5;
            this.qtdLinhas = this.graficoColunaModel5.length - 1;
        }

    }



    validaObjeto(obj: Array<GraficoImagemPosicionamento>) {
        if (obj != null && obj != undefined && obj.length > 0) {
            return true;
        }

        return false;
    }


}