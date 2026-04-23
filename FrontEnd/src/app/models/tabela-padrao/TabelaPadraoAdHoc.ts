export class TabelaAdHoc {
    CodMarca: number;
    DescMarca: string;
    CorSiteMarca: string;
    Opcao: number;
    Texto: string;
    Base: number;
    Perc1: number;
    Perc2: number;
    Perc3: number;
    Perc4: number;
    Perc5: number;
    Perc6: number;
    Perc7: number;
    Perc8: number;
    Perc9: number;
    Perc10: number;
}

export class TabelaAdHocAtributo {
    Atributo: number;
    DescAtributo: string;
}

export class TabelaPadraoAdHoc {
    Titulos: TabelaAdHocAtributo[];
    Dados: TabelaAdHoc[];
}