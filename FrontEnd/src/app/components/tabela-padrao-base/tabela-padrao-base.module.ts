import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { TabelaPadraoBaseComponent } from './tabela-padrao-base.component';
import { NgxPaginationModule } from 'ngx-pagination';
import { NgSelectModule } from '@ng-select/ng-select';
import { FormsModule } from '@angular/forms';


@NgModule({
    declarations: [
      TabelaPadraoBaseComponent
    ],
    imports: [
        CommonModule,
        NgxPaginationModule,
        NgSelectModule,
        FormsModule 
    ],
    exports: [TabelaPadraoBaseComponent]
})

export class TabelaPadraoBaseModule { }
