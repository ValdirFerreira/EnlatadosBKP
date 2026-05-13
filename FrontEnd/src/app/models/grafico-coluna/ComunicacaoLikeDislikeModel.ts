export class ComunicacaoLikeDislikeModel {

    CodMarca: number;
    DescMarca: string;

    BaseAbs: number;

    // Gostei muito
    PercGostei: number;
    TesteSIGGostei: string;

    // Gostei um pouco
    PercGosteiPouco: number;
    TesteSIGGosteiPouco: string;

    // Não gostei nem desgostei
    PercNenhum: number;
    TesteSIGNenhum: string;

    // Não gostei muito
    PercNaoGostei: number;
    TesteSigNaoGostei: string;

    // Não gostei nada
    PercNaoGosteiPouco: number;
    TesteSigNaoGosteiPouco: string;

    // T2B
    PercT2B: number;
    TesteSigT2B: string;

    // Base mínima
    BaseMinima: string;
}