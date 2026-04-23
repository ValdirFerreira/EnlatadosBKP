import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core'
import { Router } from '@angular/router';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';

import { AuthService } from 'src/app/services/auth.service';
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';
import { MenuService } from 'src/app/services/menu.service';

@Component({
  selector: 'select-checkbox-atributo',
  templateUrl: './select-checkbox-atributo.component.html',
  styleUrls: ['./select-checkbox-atributo.component.scss'],
})
export class SelectCheckboxAtributoComponent implements OnInit {

  @Output('ModelAtributo') ModelAtributo: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

  @Input() IndexCombo: number = 0;

  @Input('tipoBia') tipoBia: boolean = true;

  @Input('tipoMarca') tipoMarca: boolean = false;

  descricaoAtributo: string = "";

  constructor(public router: Router,
    private authService: AuthService, public filtroService: FiltroGlobalService,
    public menuService: MenuService
  ) { }

  public listaAtributo: Array<PadraoComboFiltro>;
  // //MODEL

  itemSelect:number = 1;

  activeCheckbox: boolean = false;

  ngOnInit(): void {
    // if (!this.tipoMarca) {
    //   this.FiltroAtributos();
    //   alert("atri")
    // }
    // else {
    //   this.FiltroMarcas();
    //   alert("mar")
    // }

    this.FiltroAtributos();

    localStorage.removeItem('ultimoSelecionado');
  }


  // ngOnChanges() {
  //   alert(this.tipoBia)
  // }

  onAtributoChanged(atributo: PadraoComboFiltro): void {
    this.ModelAtributo.emit(atributo);

        localStorage.setItem('ultimoSelecionado', atributo.IdItem.toString());
    this.descricaoAtributo = atributo.DescItem.length > 50 ? atributo.DescItem.substring(0, 50) + "..." : atributo.DescItem;
    this.activeCheckbox = false;
  }

  openContainerAtributo() {
    if (this.activeCheckbox)
      this.activeCheckbox = false;
    else
      this.activeCheckbox = true;
  }


  FiltroAtributos() {
    var idFiltro = this.filtroService.ModelTarget ? this.filtroService.ModelTarget.IdItem : 1

    debugger

    this.filtroService.FiltroAtributos(idFiltro)
      .subscribe((response: Array<PadraoComboFiltro>) => {
        debugger
        this.listaAtributo = response;
       const saved = localStorage.getItem('ultimoSelecionado');
        if (saved) {
          this.itemSelect = +saved; // converte para número
        } else
          this.itemSelect = response[0].IdItem
          
        this.descricaoAtributo = this.listaAtributo[0].DescItem.length > 50 ? this.listaAtributo[0].DescItem.substring(0, 50) + "..." : this.listaAtributo[0].DescItem;
      }, (error) => console.error(error),
        () => {
        }
      )
  }


  FiltroMarcas() {
    this.filtroService.FiltroMarcas(this.filtroService.ModelRegiao)
      .subscribe((response: Array<PadraoComboFiltro>) => {
        this.listaAtributo = response;
        this.descricaoAtributo = this.listaAtributo[0].DescItem.length > 50 ? this.listaAtributo[0].DescItem.substring(0, 50) + "..." : this.listaAtributo[0].DescItem;
      }, (error) => console.error(error),
        () => {
        }
      )
  }



}

