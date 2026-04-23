using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entities.GraficoComunicacaoDiagnostico
{
    public class GraficoComunicacaoDiagnostico
    {
        public string Descricao { get; set; }
        public decimal Perc { get; set; }
        public decimal PercNorma { get; set; }
        public decimal GAP { get; set; }
        public string TesteSig { get; set; }
        public string BaseMinima { get; set; }
        public decimal Base { get; set; }
    }
}
