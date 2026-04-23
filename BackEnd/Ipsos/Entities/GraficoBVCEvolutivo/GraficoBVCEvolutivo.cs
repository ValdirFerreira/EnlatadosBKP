using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entities.GraficoBVCEvolutivo
{
    public class GraficoBVCEvolutivo
    {
        public int CodMarca { get; set; }
        public string DescMarca { get; set; }
        public string CorSiteMarca { get; set; }
        public decimal Perc1 { get; set; }
        public decimal BasePerc1 { get; set; }

        public string BaseMinimaPerc1 { get; set; }
    }

    public class GraficoBVCEvolutivoFull
    {

        public GraficoBVCEvolutivoFull()
        {
            GraficoBVCEvolutivo1 = new List<GraficoBVCEvolutivo>();
            GraficoBVCEvolutivo2 = new List<GraficoBVCEvolutivo>();
            GraficoBVCEvolutivo3 = new List<GraficoBVCEvolutivo>();
            GraficoBVCEvolutivo4 = new List<GraficoBVCEvolutivo>();
        }

        public List<GraficoBVCEvolutivo> GraficoBVCEvolutivo1 { get; set; }
        public List<GraficoBVCEvolutivo> GraficoBVCEvolutivo2 { get; set; }
        public List<GraficoBVCEvolutivo> GraficoBVCEvolutivo3 { get; set; }
        public List<GraficoBVCEvolutivo> GraficoBVCEvolutivo4 { get; set; }

    }
}
