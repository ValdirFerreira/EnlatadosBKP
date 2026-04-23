import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { TabelaPadraoComponent } from './tabela-padrao.component';
import { NgxPaginationModule } from 'ngx-pagination';
import { NgSelectModule } from '@ng-select/ng-select';
import { FormsModule } from '@angular/forms';


@NgModule({
    declarations: [
      TabelaPadraoComponent
    ],
    imports: [
        CommonModule,
        NgxPaginationModule,
        NgSelectModule,
        FormsModule 
    ],
    exports: [TabelaPadraoComponent]
})

export class TabelaPadraoModule { }
