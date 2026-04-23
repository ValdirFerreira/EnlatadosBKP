import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core'
import { Router } from '@angular/router';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';

import { AuthService } from 'src/app/services/auth.service';
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';
import { MenuService } from 'src/app/services/menu.service';

@Component({
  selector: 'select-checkbox-adHoc',
  templateUrl: './select-checkbox-adHoc.component.html',
  styleUrls: ['./select-checkbox-adHoc.component.scss'],
})
export class SelectCheckboxAdHocComponent implements OnInit {

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


  activeCheckbox: boolean = false;

  ngOnInit(): void {
      this.FiltroAtributos();  

  }


  // ngOnChanges() {
  //   alert(this.tipoBia)
  // }

  onAtributoChanged(atributo: PadraoComboFiltro): void {
    this.ModelAtributo.emit(atributo);

    // this.descricaoAtributo = atributo.DescItem.length > 50 ? atributo.DescItem.substring(0, 50) + "..." : atributo.DescItem;
    this.descricaoAtributo =  atributo.DescItem;
    this.activeCheckbox = false;
  }

  openContainerAtributo() {
    if (this.activeCheckbox)
      this.activeCheckbox = false;
    else
      this.activeCheckbox = true;
  }


  FiltroAtributos() {
    this.filtroService.FiltroAdHoc()
      .subscribe((response: Array<PadraoComboFiltro>) => {
        this.listaAtributo = response;
        this.descricaoAtributo =  this.listaAtributo[0].DescItem;
        // this.descricaoAtributo = this.listaAtributo[0].DescItem.length > 50 ? this.listaAtributo[0].DescItem.substring(0, 50) + "..." : this.listaAtributo[0].DescItem;
      }, (error) => console.error(error),
        () => {
        }
      )
  }

}

