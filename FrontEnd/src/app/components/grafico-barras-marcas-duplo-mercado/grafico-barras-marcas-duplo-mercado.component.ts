import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';
import { GraficoColunaModel } from 'src/app/models/grafico-coluna/grafico-coluna';
import { GraficoImagemPosicionamento } from 'src/app/models/grafico-Imagem-posicionamento/GraficoImagemPosicionamento';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';


@Component({
    selector: 'grafico-barras-marcas-duplo-mercado',
    templateUrl: './grafico-barras-marcas-duplo-mercado.component.html',
    styleUrls: ['./grafico-barras-marcas-duplo-mercado.component.scss']
})


export class GraficoBarrasMarcasDuploMercadoComponent implements OnInit {

    @Input() graficoColunaModel10: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel20: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel30: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel40: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel50: Array<GraficoImagemPosicionamento>;

    @Input() graficoColunaModel60: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel70: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel80: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel90: Array<GraficoImagemPosicionamento>;
    @Input() graficoColunaModel100: Array<GraficoImagemPosicionamento>;

    @Output('marcaColuna1') marcaColuna1: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna2') marcaColuna2: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna3') marcaColuna3: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna4') marcaColuna4: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('marcaColuna5') marcaColuna5: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();


    @Output('ondaColuna1') ondaColuna1: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna2') ondaColuna2: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna3') ondaColuna3: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna4') ondaColuna4: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna5') ondaColuna5: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

    @Output('ondaColuna6') ondaColuna6: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna7') ondaColuna7: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna8') ondaColuna8: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna9') ondaColuna9: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
    @Output('ondaColuna10') ondaColuna10: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();


    onda1: PadraoComboFiltro;
    onda2: PadraoComboFiltro;
    onda3: PadraoComboFiltro;
    onda4: PadraoComboFiltro;
    onda5: PadraoComboFiltro;

    onda6: PadraoComboFiltro;
    onda7: PadraoComboFiltro;
    onda8: PadraoComboFiltro;
    onda9: PadraoComboFiltro;
    onda10: PadraoComboFiltro;
    qtdLinhas:number = 0;


 graficoColunaDescricao: Array<GraficoImagemPosicionamento>;

    ngOnInit(): void {
        this.FiltroOndas();
    
    }

    ngOnChanges(changes: any) {
    this.verificaDescricao();
  
        //  this.qtdLinhas = this.graficoColunaModel10.length - 1;
    }

    constructor(
        public filtroService: FiltroGlobalService,
    ) { 

       
    }


    porcBarraPositiva(valor) {
     
        if (valor > 0 && valor < 1)
        return "4%";
        
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

    onchangeOndaColuna6(item: PadraoComboFiltro) {
        this.ondaColuna6.emit(item);
    }

    onchangeOndaColuna7(item: PadraoComboFiltro) {
        this.ondaColuna7.emit(item);
    }

    onchangeOndaColuna8(item: PadraoComboFiltro) {
        this.ondaColuna8.emit(item);
    }

    onchangeOndaColuna9(item: PadraoComboFiltro) {
        this.ondaColuna9.emit(item);
    }

    onchangeOndaColuna10(item: PadraoComboFiltro) {
        this.ondaColuna10.emit(item);
    }


    FiltroOndas() {

        if (this.filtroService.listaOnda) {
            var listaOndas = this.filtroService.listaOnda;

            this.onda1 = listaOndas[1];
            this.onda2 = listaOndas[0];
            this.onda3 = listaOndas[1];
            this.onda4 = listaOndas[0];

            this.onda5 = listaOndas[1];
            this.onda6 = listaOndas[0];
            this.onda7 = listaOndas[1];
            this.onda8 = listaOndas[0];

            this.onda9 = listaOndas[1];
            this.onda10 = listaOndas[0];
        }
        else {

            this.filtroService.FiltroOnda()
                .subscribe((response: Array<PadraoComboFiltro>) => {
                    var listaOndas = response;

                    this.onda1 = listaOndas[1];
                    this.onda2 = listaOndas[0];
                    this.onda3 = listaOndas[1];
                    this.onda4 = listaOndas[0];

                    this.onda5 = listaOndas[1];
                    this.onda6 = listaOndas[0];
                    this.onda7 = listaOndas[1];
                    this.onda8 = listaOndas[0];

                    this.onda9 = listaOndas[1];
                    this.onda10 = listaOndas[0];

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
        if (!this.graficoColunaModel10.length)
            return '';

        if (this.validaObjeto(this.graficoColunaModel10)) {
            return this.graficoColunaModel10[0].BaseMinima;
        }

        if (this.validaObjeto(this.graficoColunaModel20)) {
            return this.graficoColunaModel20[0].BaseMinima;
        }

        if (this.validaObjeto(this.graficoColunaModel30)) {
            return this.graficoColunaModel30[0].BaseMinima;
        }

        if (this.validaObjeto(this.graficoColunaModel40)){
            return this.graficoColunaModel40[0].BaseMinima;
        }

        if (this.validaObjeto(this.graficoColunaModel50 )) {
            return this.graficoColunaModel50[0].BaseMinima;
        }

        if (this.validaObjeto(this.graficoColunaModel60 )) {
            return this.graficoColunaModel60[0].BaseMinima;
        }

        if (this.validaObjeto(this.graficoColunaModel70)) {
            return this.graficoColunaModel70[0].BaseMinima;
        }

        if (this.validaObjeto(this.graficoColunaModel80 )) {
            return this.graficoColunaModel80[0].BaseMinima;
        }

        if (this.validaObjeto(this.graficoColunaModel90 )) {
            return this.graficoColunaModel90[0].BaseMinima;
        }

        if (this.validaObjeto(this.graficoColunaModel100)) {
            return this.graficoColunaModel100[0].BaseMinima;
        }

        return '';
    }


    verificaDescricao() {
 

        if (this.validaObjeto(this.graficoColunaModel10)) {
            this.graficoColunaDescricao = this.graficoColunaModel10;
            this.qtdLinhas = this.graficoColunaModel10.length - 1;
        }

        if (this.validaObjeto(this.graficoColunaModel20)) {
            this.graficoColunaDescricao = this.graficoColunaModel20;
            this.qtdLinhas = this.graficoColunaModel20.length - 1;
        }

        if (this.validaObjeto(this.graficoColunaModel30)) {
            this.graficoColunaDescricao = this.graficoColunaModel30;
            this.qtdLinhas = this.graficoColunaModel30.length - 1;
        }

        if (this.validaObjeto(this.graficoColunaModel40)){
            this.graficoColunaDescricao = this.graficoColunaModel40;
            this.qtdLinhas = this.graficoColunaModel40.length - 1;
        }

        if (this.validaObjeto(this.graficoColunaModel50 )) {
            this.graficoColunaDescricao = this.graficoColunaModel50;
            this.qtdLinhas = this.graficoColunaModel50.length - 1;
        }

        if (this.validaObjeto(this.graficoColunaModel60 )) {
            this.graficoColunaDescricao = this.graficoColunaModel60;
            this.qtdLinhas = this.graficoColunaModel60.length - 1;
        }

        if (this.validaObjeto(this.graficoColunaModel70)) {
            this.graficoColunaDescricao = this.graficoColunaModel70;
            this.qtdLinhas = this.graficoColunaModel70.length - 1;
        }

        if (this.validaObjeto(this.graficoColunaModel80 )) {
            this.graficoColunaDescricao = this.graficoColunaModel80;
            this.qtdLinhas = this.graficoColunaModel80.length - 1;
        }

        if (this.validaObjeto(this.graficoColunaModel90 )) {
            this.graficoColunaDescricao = this.graficoColunaModel90;
            this.qtdLinhas = this.graficoColunaModel90.length - 1;
        }

        if (this.validaObjeto(this.graficoColunaModel100)) {
            this.graficoColunaDescricao = this.graficoColunaModel100;
            this.qtdLinhas = this.graficoColunaModel100.length - 1;
        }

 
    }


    
    validaObjeto(obj: Array<GraficoImagemPosicionamento>) {
        if (obj != null && obj != undefined && obj.length > 0) {
            return true;
        }

        return false;
    }




}