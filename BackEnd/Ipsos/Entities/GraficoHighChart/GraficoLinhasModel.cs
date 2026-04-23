using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entities.GraficoHighChart
{
    public class GraficoLinhasModel
    {
        public List<string> Periodos { get; set; }
        public List<GraficoLinhasHighChartModel> Grafico { get; set; }
        public List<string> Cores { get; set; }

        public List<string> Bases { get; set; }
    }
    public class GraficoLinhasHighChartModel
    {
        public GraficoLinhasHighChartModel()
        {
            data = new List<DataModel>();
        }

        public string type { get; set; }
        public string name { get; set; }
        public List<DataModel> data { get; set; }
        public object marker { get; set; }

        public int tipoDados { get; set; }

    }

    public class DataModel
    {
        public string name { get; set; }
        public decimal y { get; set; }
        public string periodo { get; set; }
        public decimal valorbase { get; set; }
        public string sig { get; set; }
        public int tipoDados { get; set; }
        public decimal media { get; set; }

        public string baseminima { get; set; }

        public decimal BaseAnt { get; set; }

        public decimal MediaAnt { get; set; }

        public string BaseMinimaAnt { get; set; }

    }

    public class GraficoLinhasRetornoProcedurePadrao
    {
        public string DescPeriodo { get; set; }
        public string Descricao { get; set; }
        public decimal Valor { get; set; }
        public decimal Base { get; set; }
        public int TipoDados { get; set; }

    }
}
