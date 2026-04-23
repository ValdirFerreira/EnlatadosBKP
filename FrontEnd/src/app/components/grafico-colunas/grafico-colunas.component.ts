import { Component, Input, OnInit } from '@angular/core';
import { GraficoColunaModel } from 'src/app/models/grafico-coluna/grafico-coluna';


@Component({
    selector: 'grafico-colunas',
    templateUrl: './grafico-colunas.component.html',
    styleUrls: ['./grafico-colunas.component.scss']
})
export class GraficoColunasComponent implements OnInit {
    
    @Input() graficoColunaModel: Array<any>;
 

    ngOnInit(): void { 
    } 
    
    ajusteDeColuna1(Promotores: any, Neutros: any, Detratores: any, nps: any) {
        var retorno;
        if (Neutros >= 15) {
            retorno = Promotores;
        } else retorno = '85%'

        if (Detratores >= 10) {
        } else retorno = '90%'

        if (Neutros >= 15 && Detratores >= 10) {
            retorno = Promotores
        } else retorno = '75'

        var total = Neutros + Detratores;
        var total3colunas = Neutros + Detratores + Promotores;

        if (total3colunas < 100) {
            retorno = 100 - total
        } else {
            retorno = Promotores;
        }

        if (Neutros == 0 && Detratores == 0) {
            retorno = '100'
        }

        if (Neutros > 0 && Detratores == 0 || Neutros == 0 && Detratores > 0) {
            retorno = '50';
        }

        if (nps == 0) {
            if (Neutros > 0 && Detratores > 0 || Neutros > 0 && Detratores > 0) {
                retorno = '50';
            }
        }
        return retorno + '%';
    }

    ajusteDeColuna2(Promotores: any, Neutros: any, Detratores: any, nps: any): string {
        var retorno;

        if (Neutros >= 15) {
            retorno = Neutros;
        } else {
            retorno = '15';
        }

        if (Promotores == 0 && Detratores == 0) {
            retorno = '100';
        }

        if (Promotores > 0 && Detratores == 0 || Promotores == 0 && Detratores > 0) {
            retorno = '50';
        }

        if (nps == 0) {
            if (Promotores > 0 && Detratores > 0 || Promotores > 0 && Detratores > 0) {
                retorno = '25';
            }
        }

        return retorno + '%';
    }

    ajusteDeColuna3(Promotores: any, Neutros: any, Detratores: any, nps: any): string {
        var retorno;

        if (Detratores >= 10) {
            retorno = Detratores;
        } else {
            retorno = '10';
        }

        if (Promotores == 0 && Neutros == 0) {
            retorno = '100';
        }

        if (Promotores > 0 && Neutros == 0 || Promotores == 0 && Neutros > 0) {
            retorno = '50';
        }

        if (nps == 0) {
            if (Promotores > 0 && Neutros > 0 || Promotores > 0 && Neutros > 0) {
                retorno = '25';
            }
        }

        return retorno + '%';
    }

}