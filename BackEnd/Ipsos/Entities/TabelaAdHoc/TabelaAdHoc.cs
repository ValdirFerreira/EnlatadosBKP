using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entities.TabelaAdHoc
{
    public class TabelaAdHoc
    {
        public int CodMarca { get; set; }
        public string DescMarca { get; set; }
        public string CorSiteMarca { get; set; }
        public int Opcao { get; set; }
        public string Texto { get; set; }

        public decimal Base { get; set; }

        public decimal Perc1 { get; set; }
        public decimal Perc2 { get; set; }
        public decimal Perc3 { get; set; }
        public decimal Perc4 { get; set; }
        public decimal Perc5 { get; set; }
        public decimal Perc6 { get; set; }
        public decimal Perc7 { get; set; }
        public decimal Perc8 { get; set; }
        public decimal Perc9 { get; set; }
        public decimal Perc10 { get; set; }
    }

    public class TabelaAdHocAtributo
    {
        public int Atributo { get; set; }
        public string DescAtributo { get; set; }
    }


    public class TabelaPadraoAdHoc
    {

        public TabelaPadraoAdHoc()
        {
            Titulos = new List<TabelaAdHocAtributo>();
            Dados = new List<TabelaAdHoc>();
        }

        public List<TabelaAdHocAtributo> Titulos { get; set; }

        public List<TabelaAdHoc> Dados { get; set; }
    }


}
