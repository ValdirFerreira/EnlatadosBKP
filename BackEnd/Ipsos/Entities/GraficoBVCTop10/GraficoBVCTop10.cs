using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entities.GraficoBVCTop10
{
    public class GraficoBVCTop10
    {
        public int CodMarca { get; set; }
        public string DescMarca { get; set; }
        public string CorSiteMarca { get; set; }
        public decimal ShareDesejo { get; set; }
        public decimal Efeitos { get; set; }
        public decimal Equity { get; set; }

        public decimal BaseShareDesejo { get; set; }
        public decimal BaseEfeitos { get; set; }
        public decimal BaseEquity { get; set; }

        public string BaseMinimaShareDesejo { get; set; }
        public string BaseMinimaEfeitos { get; set; }
        public string BaseMinimaEquity { get; set; }
    }
}
