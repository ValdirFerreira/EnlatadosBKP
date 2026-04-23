using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entities.GraficoImagemPosicionamento
{
    public class GraficoImagemPosicionamento
    {
        public int CodMarca { get; set; }
        public string DescMarca { get; set; }
        public string CorSiteMarca { get; set; }
        public decimal Resposta { get; set; }
        public string Descricao { get; set; }
        public decimal PercAtual { get; set; }
        public decimal PercAnterior { get; set; }
        public string TestSig { get; set; }
        public decimal Base { get; set; }
        public decimal Media { get; set; }
        public string Negrito { get; set; }
        public string BaseMinima { get; set; }
    }


    public class GraficoImagemPosicionamentoFullLoad
    {

        public GraficoImagemPosicionamentoFullLoad()
        {
            GraficoImagemPosicionamento1 = new List<GraficoImagemPosicionamento>();
            GraficoImagemPosicionamento2 = new List<GraficoImagemPosicionamento>();
            GraficoImagemPosicionamento1_3 = new List<GraficoImagemPosicionamento>();

            GraficoImagemPosicionamento3 = new List<GraficoImagemPosicionamento>();
            GraficoImagemPosicionamento4 = new List<GraficoImagemPosicionamento>();
            GraficoImagemPosicionamento3_3 = new List<GraficoImagemPosicionamento>();

            GraficoImagemPosicionamento5 = new List<GraficoImagemPosicionamento>();
            GraficoImagemPosicionamento6 = new List<GraficoImagemPosicionamento>();
            GraficoImagemPosicionamento5_3 = new List<GraficoImagemPosicionamento>();

            GraficoImagemPosicionamento7 = new List<GraficoImagemPosicionamento>();
            GraficoImagemPosicionamento8 = new List<GraficoImagemPosicionamento>();
            GraficoImagemPosicionamento7_3 = new List<GraficoImagemPosicionamento>();

            GraficoImagemPosicionamento9 = new List<GraficoImagemPosicionamento>();
            GraficoImagemPosicionamento10 = new List<GraficoImagemPosicionamento>();
            GraficoImagemPosicionamento9_3 = new List<GraficoImagemPosicionamento>();
        }


        public List<GraficoImagemPosicionamento> GraficoImagemPosicionamento1 { get; set; }
        public List<GraficoImagemPosicionamento> GraficoImagemPosicionamento2 { get; set; }
        public List<GraficoImagemPosicionamento> GraficoImagemPosicionamento1_3 { get; set; }

        public List<GraficoImagemPosicionamento> GraficoImagemPosicionamento3 { get; set; }
        public List<GraficoImagemPosicionamento> GraficoImagemPosicionamento4 { get; set; }
        public List<GraficoImagemPosicionamento> GraficoImagemPosicionamento3_3 { get; set; }

        public List<GraficoImagemPosicionamento> GraficoImagemPosicionamento5 { get; set; }
        public List<GraficoImagemPosicionamento> GraficoImagemPosicionamento6 { get; set; }
        public List<GraficoImagemPosicionamento> GraficoImagemPosicionamento5_3 { get; set; }

        public List<GraficoImagemPosicionamento> GraficoImagemPosicionamento7 { get; set; }
        public List<GraficoImagemPosicionamento> GraficoImagemPosicionamento8 { get; set; }
        public List<GraficoImagemPosicionamento> GraficoImagemPosicionamento7_3 { get; set; }

        public List<GraficoImagemPosicionamento> GraficoImagemPosicionamento9 { get; set; }
        public List<GraficoImagemPosicionamento> GraficoImagemPosicionamento10 { get; set; }

        public List<GraficoImagemPosicionamento> GraficoImagemPosicionamento9_3 { get; set; }

    }

}
