using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entities.GraficoColunas
{
    public class GraficoColunas
    {
        public int CodMarca { get; set; }
        public string DescMarca { get; set; }
        public decimal BaseAbs { get; set; }
        public decimal PercTOM { get; set; }
        public string TesteSIGTOM { get; set; }
        public decimal PercOM { get; set; }
        public string TesteSIGOM { get; set; }
        public decimal PercPrompeted { get; set; }
        public string TesteSIGPrompeted { get; set; }
        public decimal PercTotal { get; set; }
        public string TesteSigTotal { get; set; }
        public decimal PercOutros { get; set; }
        public string TesteSigOutros { get; set; }
        public string BaseMinima { get; set; }

    }


    public class GraficoColunasFullLoad
    {
        public GraficoColunasFullLoad()
        {
            GraficoColunas1 = new GraficoColunas();
            GraficoColunas2 = new GraficoColunas();
            GraficoColunas3 = new GraficoColunas();
            GraficoColunas4 = new GraficoColunas();
            GraficoColunas5 = new GraficoColunas();
        }

        public GraficoColunas GraficoColunas1 { get; set; }
        public GraficoColunas GraficoColunas2 { get; set; }
        public GraficoColunas GraficoColunas3 { get; set; }
        public GraficoColunas GraficoColunas4 { get; set; }
        public GraficoColunas GraficoColunas5 { get; set; }
    }




}
