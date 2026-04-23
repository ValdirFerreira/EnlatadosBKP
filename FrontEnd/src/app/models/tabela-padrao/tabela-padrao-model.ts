export class TabelaPadraoModel {
    Titulos: Array<string>;
    Periodos: Array<string>;
    Linhas: Array<ColunasTabelaPadraoModel>;
    TotalRegistro?: number;
}

export class ColunasTabelaPadraoModel {
    Descricao: string;
    Coluna: Array<string>;
}
