import { Component, Input, OnInit } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { GraficoComunicacaoDiagnostico } from 'src/app/models/GraficoComunicacaoDiagnostico/GraficoComunicacaoDiagnostico';

@Component({
    selector: 'grafico-diagnostico-propaganda',
    templateUrl: './grafico-diagnostico-propaganda.component.html',
    styleUrls: ['./grafico-diagnostico-propaganda.component.scss']
})
export class GraficoDiagnosticoPropagandaComponent implements OnInit {
    
    @Input() graficoComunicacaoDiagnostico: Array<GraficoComunicacaoDiagnostico>;
 
    constructor(
        private translate: TranslateService,
    ) { }

    ngOnInit(): void { 
    } 
    

    calcularPorcentagemColuna(valor): any {

        if (valor > 100)
          return 100;

          if(valor == 0)
          return 0.04;
          
        return valor;
      }


      verificaMsgBase() {

        var baseMinima = '';
    
        this.graficoComunicacaoDiagnostico.forEach(item => {
          if (item.BaseMinima != '')
            baseMinima = item.BaseMinima;
        })
    
        return baseMinima;
      }


      validaSig(valor: string) {

        if (valor == 'MAIOR')
          return 'sig-positive';
    
        if (valor == 'MENOR')
          return 'sig-negative';
    
        if (valor == 'IGUAL')
          return 'sig-vazio';
    
        return 'sig-vazio';
      }
  
}