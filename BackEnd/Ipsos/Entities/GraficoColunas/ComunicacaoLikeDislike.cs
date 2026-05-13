using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entities.GraficoColunas
{
    public class ComunicacaoLikeDislike
    {
        public int CodMarca { get; set; }
        public string DescMarca { get; set; }

        public decimal BaseAbs { get; set; }

        // Gostei muito
        public decimal PercGostei { get; set; }
        public string TesteSIGGostei { get; set; }

        // Gostei um pouco
        public decimal PercGosteiPouco { get; set; }
        public string TesteSIGGosteiPouco { get; set; }

        // Não gostei nem desgostei
        public decimal PercNenhum { get; set; }
        public string TesteSIGNenhum { get; set; }

        // Não gostei muito
        public decimal PercNaoGostei { get; set; }
        public string TesteSigNaoGostei { get; set; }

        // Não gostei nada
        public decimal PercNaoGosteiPouco { get; set; }
        public string TesteSigNaoGosteiPouco { get; set; }

        // T2B
        public decimal PercT2B { get; set; }
        public string TesteSigT2B { get; set; }

        // Base mínima
        public string BaseMinima { get; set; }
    }

    public class ComunicacaoLikeDislikeFullLoad
    {
        public ComunicacaoLikeDislikeFullLoad()
        {
            ComunicacaoLikeDislike1 = new ComunicacaoLikeDislike();
            ComunicacaoLikeDislike2 = new ComunicacaoLikeDislike();
            ComunicacaoLikeDislike3 = new ComunicacaoLikeDislike();
            ComunicacaoLikeDislike4 = new ComunicacaoLikeDislike();
            ComunicacaoLikeDislike5 = new ComunicacaoLikeDislike();
        }

        public ComunicacaoLikeDislike ComunicacaoLikeDislike1 { get; set; }
        public ComunicacaoLikeDislike ComunicacaoLikeDislike2 { get; set; }
        public ComunicacaoLikeDislike ComunicacaoLikeDislike3 { get; set; }
        public ComunicacaoLikeDislike ComunicacaoLikeDislike4 { get; set; }
        public ComunicacaoLikeDislike ComunicacaoLikeDislike5 { get; set; }
    }
}