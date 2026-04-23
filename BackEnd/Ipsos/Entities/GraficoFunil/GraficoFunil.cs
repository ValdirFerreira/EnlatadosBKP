using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entities.GraficoFunil
{
    public class GraficoFunil
    {
        public int CodMarca { get; set; }
        public string DescMarca { get; set; }
        public string CorSiteMarca { get; set; }
        public string PeriodoAnt { get; set; }
        public decimal BaseAbsAtual { get; set; }
        public decimal BaseAbsAnterior { get; set; }
        public decimal ConhecimentoAnterior { get; set; }
        public decimal ConhecimentoAtual { get; set; }
        public string ConhecimentoTesteAtual { get; set; }
        public decimal ConhecimentoADV { get; set; }
        public decimal ConsideracaoAnterior { get; set; }
        public decimal ConsideracaoAtual { get; set; }
        public string ConsideracaoTesteAtual { get; set; }
        public decimal ConsideracaoADV { get; set; }
        public decimal UsoAnterior { get; set; }
        public decimal UsoAtual { get; set; }
        public string UsoTesteAtual { get; set; }
        public decimal UsoADV { get; set; }
        public decimal PreferenciaAnterior { get; set; }
        public decimal PreferenciaAtual { get; set; }
        public string PreferenciaTesteAtual { get; set; }
        public decimal PreferenciaADV { get; set; }

        public decimal LoyaltyAnterior { get; set; }
        public decimal LoyaltyAtual { get; set; }
        public string LoyaltyTesteAtual { get; set; }
        public decimal LoyaltyADV { get; set; }

        public string BaseMinimaAtual { get; set; }

        public string BaseMinimaAnterior { get; set; }
    }



    public class GraficoFunilFullLoad
    {

        public GraficoFunilFullLoad()
        {
            GraficoFunil1 = new GraficoFunil();
            GraficoFunil2 = new GraficoFunil();
            GraficoFunil3 = new GraficoFunil();
            GraficoFunil4 = new GraficoFunil();
            GraficoFunil5 = new GraficoFunil();
            GraficoFunil6 = new GraficoFunil();
            GraficoFunil7 = new GraficoFunil();
            GraficoFunil8 = new GraficoFunil();
           
        }

        public GraficoFunil GraficoFunil1 { get; set; }
        public GraficoFunil GraficoFunil2 { get; set; }
        public GraficoFunil GraficoFunil3 { get; set; }
        public GraficoFunil GraficoFunil4 { get; set; }
        public GraficoFunil GraficoFunil5 { get; set; }
        public GraficoFunil GraficoFunil6 { get; set; }
        public GraficoFunil GraficoFunil7 { get; set; }
        public GraficoFunil GraficoFunil8 { get; set; }
    }


}
