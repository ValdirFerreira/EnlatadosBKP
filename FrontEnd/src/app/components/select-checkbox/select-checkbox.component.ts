import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core'
import { Router } from '@angular/router';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';

import { AuthService } from 'src/app/services/auth.service';
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';
import { MenuService } from 'src/app/services/menu.service';

@Component({
  selector: 'select-checkbox',
  templateUrl: './select-checkbox.component.html',
  styleUrls: ['./select-checkbox.component.scss'],
})
export class SelectCheckboxComponent implements OnInit {

  @Output('ModeslMarca') ModeslMarca: EventEmitter<Array<PadraoComboFiltro>> = new EventEmitter<Array<PadraoComboFiltro>>();

  @Input() IndexCombo: number = 0;


  constructor(public router: Router,
    private authService: AuthService, public filtroService: FiltroGlobalService,
    public menuService: MenuService
  ) { }

  public listaMarcas: Array<PadraoComboFiltro>;
  // //MODEL
  public ModelMarcaSelecionadas = new Array<PadraoComboFiltro>();

  maxIdItem:number=0;
  minIdItem:number=0;


  activeCheckbox: boolean = false;

  ngOnInit(): void {
    this.FiltroMarcas();
  }

  onMarcaChanged(marca: PadraoComboFiltro): void {

    if (this.ModelMarcaSelecionadas.filter(s => s.IdItem == marca.IdItem).length > 0) {
      const index = this.ModelMarcaSelecionadas.indexOf(marca);
      this.ModelMarcaSelecionadas.splice(index, 1);
    }
    else
      this.ModelMarcaSelecionadas.push(marca);

  }

  openContainerMarcas() {
    if (this.activeCheckbox)
      this.activeCheckbox = false;
    else
      this.activeCheckbox = true;
  }

  FiltroMarcas() {
    this.filtroService.FiltroMarcas(this.filtroService.ModelRegiao)
      .subscribe((response: Array<PadraoComboFiltro>) => {
       // this.listaMarcas = response;

        var list =[]
        var cont =1;
        response.forEach(x=>
          {
            var item = x;
            item.Ordem =cont;
            list.push(item)
            cont++;
          })

          this.listaMarcas =list;

        this.ModelMarcaSelecionadas.push(this.listaMarcas[0])
        this.ModelMarcaSelecionadas.push(this.listaMarcas[1])
        this.ModelMarcaSelecionadas.push(this.listaMarcas[2])
        this.ModelMarcaSelecionadas.push(this.listaMarcas[3])

        this.minIdItem = this.listaMarcas[0].IdItem;
        this.maxIdItem = this.listaMarcas[3].IdItem;

      }, (error) => console.error(error),
        () => {
        }
      )
  }


  ExecutarEmitFiltrar() {

    if (this.ModelMarcaSelecionadas.length < 1) {
      this.listaMarcas = [];
      this.filtroService.FiltroMarcas(this.filtroService.ModelRegiao)
        .subscribe((response: Array<PadraoComboFiltro>) => {
          this.listaMarcas = response;

          this.ModelMarcaSelecionadas.push(this.listaMarcas[0])
          this.ModelMarcaSelecionadas.push(this.listaMarcas[1])
          this.ModelMarcaSelecionadas.push(this.listaMarcas[2])
          this.ModelMarcaSelecionadas.push(this.listaMarcas[3])

          this.ModeslMarca.emit(this.ModelMarcaSelecionadas);
          this.activeCheckbox = false;

        }, (error) => console.error(error),
          () => {
          }
        )
    }
    else {

      this.ModeslMarca.emit(this.ModelMarcaSelecionadas);
      this.activeCheckbox = false;
    }
  }


}

