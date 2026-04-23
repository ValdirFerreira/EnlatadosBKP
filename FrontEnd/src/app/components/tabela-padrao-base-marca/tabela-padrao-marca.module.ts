import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { TabelaPadraoMarcaComponent } from './tabela-padrao-marca.component';
import { NgxPaginationModule } from 'ngx-pagination';
import { NgSelectModule } from '@ng-select/ng-select';
import { FormsModule } from '@angular/forms';


@NgModule({
    declarations: [
      TabelaPadraoMarcaComponent
    ],
    imports: [
        CommonModule,
        NgxPaginationModule,
        NgSelectModule,
        FormsModule 
    ],
    exports: [TabelaPadraoMarcaComponent]
})

export class TabelaPadraoMarcaModule { }
