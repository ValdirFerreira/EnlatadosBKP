import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';
import { PadraoComboFiltro } from 'src/app/models/padrao-combo-filtro/padrao-combo-filtro';
import { TabelaPadraoModel } from 'src/app/models/tabela-padrao/tabela-padrao-model';
import { TabelaPadraoAdHoc } from 'src/app/models/tabela-padrao/TabelaPadraoAdHoc';
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';


@Component({
  selector: 'tabela-padrao-marca',
  templateUrl: './tabela-padrao-marca.component.html',
  styleUrls: ['./tabela-padrao-marca.component.scss']
})
export class TabelaPadraoMarcaComponent implements OnInit {

  @Input('grafico') tabela: TabelaPadraoAdHoc;
  @Input() colunas: Array<string>;
  @Input() paginacao: boolean = true;
  @Input() itemsPorPagina: number = 10;
  @Input() visualizacaoPorMedia: boolean = false;
  @Output('marcaGrafico') marcaGrafico: EventEmitter<PadraoComboFiltro> = new EventEmitter<PadraoComboFiltro>();
  ativaTela: boolean = false;

  @Input()  listaMarcas: Array<PadraoComboFiltro>;
  
  descInicial: string = "Nestlé";
  public ModelMarcas: PadraoComboFiltro;

  constructor(public filtroService: FiltroGlobalService,) { }

  countLine: number = 1;
  page: number = 1;
  config: any;

  id: string;
  ngOnInit(): void {
    this.id = this.gerarIdParaConfigDePaginacao();
    this.config = {
      id: this.id,
      currentPage: this.page,
      itemsPerPage: this.itemsPorPagina
    };

    this.ativaTela = true;
  }

  ngOnChanges() {
    this.page = 1;
    if (this.config) {
      this.config.currentPage = this.page;
    }
  }

  gerarIdParaConfigDePaginacao() {
    var result = '';
    var characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    var charactersLength = characters.length;
    for (var i = 0; i < 10; i++) {
      result += characters.charAt(Math.floor(Math.random() *
        charactersLength));
    }
    return result;
  }

  // carregaTabela() {
  //   var retorno = new TabelaPadraoModel();
  //   var titulos = [];
  //   for (let index = 0; index < 20; index++) {
  //     titulos.push("Titulo Titulo Titulo" + index);
  //   }

  //   retorno.Titulos = titulos;


  //   var linhas = [];
  //   // Linha
  //   for (let linha = 0; linha < 25; linha++) {

  //     var row = new ColunasTabelaPadraoModel();
  //     var Coluna = [];

  //     for (let coluna = 0; coluna < 20; coluna++) {
  //       Coluna.push("DadoDado DadoDado - " + coluna);
  //     }

  //     row.Coluna = Coluna;
  //     linhas.push(row);
  //   }
  //   retorno.Linhas = linhas;

  //   this.tabela = retorno;

  //   this.ativaTela = true;
  // }


  mudarPage(event: any) {
    this.page = event;
    this.config.currentPage = this.page;
  }

  public mudarItemsPorPage() {
    this.page = 1
    this.config.currentPage = this.page;
    this.config.itemsPerPage = this.itemsPorPagina;
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
    //     if (this.filtroService.ModelDenominators && this.filtroService.ModelDenominators.IdItem > 1) {
    //         this.descInicial = this.filtroService.listaMarcas[0].DescItem;
    //         return "assets/marcas/13500.svg";
    //     }
    //     else
    //     return "assets/marcas/13001.svg";
    // }
}

onchangeMarcaGrafico(item: PadraoComboFiltro) {
  this.marcaGrafico.emit(item);
}

}
