import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core'
import { Router } from '@angular/router';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';

import { AuthService } from 'src/app/services/auth.service';
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';
import { MenuService } from 'src/app/services/menu.service';

@Component({
  selector: 'select-image-2',
  templateUrl: './select-image-2.component.html',
  styleUrls: ['./select-image-2.component.scss'],
})
export class SelectImage2Component implements OnInit {
  @Input('titulo') tituloPagina: string;
  @Input() pageAtual: string;
  @Input() home: boolean = false;
  // @Output() ModelMarcas: PadraoComboFiltro;
  @Output('ImageModelMarcas') ImageModelMarcas: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();

  @Input() IndexCombo: number = 0;

  filtro: boolean = false;

  descInicial: string = "Nestlé";

  constructor(public router: Router,
    private authService: AuthService, public filtroService: FiltroGlobalService,
    public menuService: MenuService
  ) { }

  public listaMarcas: Array<PadraoComboFiltro>;
  // //MODEL
  public ModelMarcas: PadraoComboFiltro;


  ngOnInit(): void {
    this.FiltroMarcas();
  }

  onGroupChanged(): void {

    this.ImageModelMarcas.emit(this.ModelMarcas);
  }



  FiltroMarcas() {
    this.filtroService.FiltroMarcas(this.filtroService.ModelRegiao)

      .subscribe((response: Array<PadraoComboFiltro>) => {
        this.listaMarcas = response;
        this.ModelMarcas = response[this.IndexCombo];

      }, (error) => console.error(error),
        () => {
        }
      )
    
  }

  montaImageMarca() {
    if (this.ModelMarcas != undefined) {
      this.descInicial = this.ModelMarcas.DescItem;

      return "assets/marcas/" + this.ModelMarcas.IdItem + ".svg";
  }
  else {
      // return "assets/marcas/1.svg";

      this.descInicial = this.filtroService.listaMarcas[0].DescItem;
      return "assets/marcas/" + this.filtroService.listaMarcas[0].IdItem + ".svg";
  }
    // else {
    //   return "assets/marcas/13001.svg";
    // }
   

  }



}

