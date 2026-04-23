import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core'
import { Router } from '@angular/router';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';

import { AuthService } from 'src/app/services/auth.service';
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';
import { MenuService } from 'src/app/services/menu.service';

@Component({
  selector: 'select-checkbox-STB',
  templateUrl: './select-checkbox-STB.component.html',
  styleUrls: ['./select-checkbox-STB.component.scss'],
})
export class SelectCheckboxSTBComponent implements OnInit {

  @Output('ModelAtributo') ModelAtributo: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

  @Input() IndexCombo: number = 0;



  descricaoAtributo: string = "";

  constructor(public router: Router,
    private authService: AuthService, public filtroService: FiltroGlobalService,
    public menuService: MenuService
  ) { }

  public listaAtributo: Array<PadraoComboFiltro>;
  // //MODEL


  activeCheckbox: boolean = false;

  ngOnInit(): void {
 
  }

  ngOnChanges() {
    this.FiltroAtributos();
  }

  onAtributoChanged(atributo: PadraoComboFiltro): void {
    this.ModelAtributo.emit(atributo);

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

    if (!this.filtroService.ModelOnda) {
      this.filtroService.FiltroOnda()
        .subscribe((response: Array<PadraoComboFiltro>) => {
          this.filtroService.listaOnda = response;
          this.filtroService.ModelOnda = response[0];

          this.carregaFiltroSTB();
        }, (error) => console.error(error),
          () => {
          }
        )
    }
    else {
      this.carregaFiltroSTB();
    }



  }

  carregaFiltroSTB() {
    this.filtroService.FiltroSTB(this.filtroService.ModelOnda)
      .subscribe((response: Array<PadraoComboFiltro>) => {
  
        this.listaAtributo = response;
        this.descricaoAtributo = this.listaAtributo[0].DescItem.length > 50 ? this.listaAtributo[0].DescItem.substring(0, 50) + "..." : this.listaAtributo[0].DescItem;
      }, (error) => console.error(error),
        () => {
        }
      )
  }



}

