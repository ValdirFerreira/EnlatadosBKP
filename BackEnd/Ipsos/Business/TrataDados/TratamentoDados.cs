using Entities.GraficoAranha;
using Entities.GraficoHighChart;
using Entities.GraficoImagemEvolutivo;
using Entities.GraficoLinhasImagem;
using Entities.TabelaAdHoc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Business.TrataDados
{
    public class TratamentoDados
    {


        private static List<GraficoLinhasImagemProcedure> GeraLinhaAnterior(List<GraficoLinhasImagemProcedure> grafico)
        {
            var dados = new List<GraficoLinhasImagemProcedure>();

            grafico.ForEach(d =>
            {
                d.TipoDados = 0; // Tipo de dados 0 LInha Marca Banco
            });
            dados.AddRange(grafico);

            grafico.ForEach(d =>
            {
                // geração dos dados anteriores
                var item = new GraficoLinhasImagemProcedure
                {
                    CodMarca = d.CodMarca,
                    DescMarca = String.Concat(d.DescMarca, "-OLD"),
                    CorSiteMarca = d.CorSiteMarca,
                    Descricao = d.Descricao,
                    GrupoAtributo = d.GrupoAtributo,
                    CorAtributo = d.CorAtributo,
                    PercAnterior = d.PercAnterior,
                    PercAtual = d.PercAnterior,
                    Resposta = d.Resposta,
                    TestSig = d.TestSig,
                    TipoDados = 1,
                    Base = d.BaseAnt,
                    Media = d.MediaAnt,
                    BaseMinima = d.BaseMinimaAnt,
                };
                dados.Add(item);
            });



            return dados;
        }

        public static GraficoLinhasModel TratarDadosGraficoLinhasImagem(List<GraficoLinhasImagemProcedure> dados)
        {

            var grafico = new List<GraficoLinhasImagemProcedure>();

            grafico = GeraLinhaAnterior(dados);

            var retorno = new GraficoLinhasModel();
            retorno.Cores = new List<string>();
            retorno.Periodos = new List<string>();
            retorno.Bases = new List<string>();
            retorno.Grafico = new List<GraficoLinhasHighChartModel>();

            if (grafico.Any())
            {

                var categorias = grafico.Select(a => new { a.DescMarca, a.TipoDados }).Distinct().ToList();


                foreach (var item in grafico.Select(a => new { a.Descricao }).Distinct().ToList())
                {
                    retorno.Periodos.Add(item.Descricao);
                }

                foreach (var item in categorias)
                {

                    var graficoModel = new GraficoLinhasHighChartModel();

                    graficoModel.name = item.DescMarca;
                    graficoModel.type = "line";
                    graficoModel.tipoDados = item.TipoDados;


                    var listaDados = grafico.Where(a => a.DescMarca.Equals(item.DescMarca)).Select(a => new
                    {
                        a.PercAtual,
                        a.Descricao,
                        a.DescMarca,
                        a.TestSig,
                        a.TipoDados,
                        a.Base,
                        a.Media,
                        a.BaseMinima,
                        a.BaseAnt,
                        a.MediaAnt,
                        a.BaseMinimaAnt,
                        a.GrupoAtributo,
                        a.CorAtributo
                    }).ToList();

                    foreach (var itemArray in listaDados)
                    {
                        graficoModel.data.Add(new DataModel
                        {
                            name = itemArray.DescMarca,
                            periodo = itemArray.Descricao,
                            valorbase = itemArray.Base,
                            y = itemArray.PercAtual,
                            //sig = itemArray.TipoDados == 1 ? "" : itemArray.TestSig
                            sig = itemArray.TestSig,
                            media = itemArray.Media,
                            baseminima = itemArray.BaseMinima,
                            BaseAnt = itemArray.BaseAnt,
                            MediaAnt = itemArray.MediaAnt,
                            BaseMinimaAnt = itemArray.BaseMinimaAnt,
                            GrupoAtributo = itemArray.GrupoAtributo,
                            CorAtributo = itemArray.CorAtributo
                        });
                    }

                    retorno.Grafico.Add(graficoModel);
                }
            }
            else
                return null;

            return retorno;
        }

        public static GraficoLinhasModel TratarDadosGraficoAranha(List<GraficoAranhaRetornoProcedure> grafico)
        {
            var retorno = new GraficoLinhasModel();
            retorno.Cores = new List<string>();
            retorno.Periodos = new List<string>();
            retorno.Bases = new List<string>();
            retorno.Grafico = new List<GraficoLinhasHighChartModel>();

            if (grafico.Any())
            {

                var categorias = grafico.Select(a => new { a.DescMarca }).Distinct().ToList();


                foreach (var item in grafico.Select(a => new { a.Descricao }).Distinct().ToList())
                {

                    string descricao = string.Concat(item.Descricao);
                    //string valorBase = string.Concat(item.Base);
                    //retorno.Bases.Add(valorBase);
                    retorno.Periodos.Add(descricao);
                }

                foreach (var item in categorias)
                {

                    var graficoModel = new GraficoLinhasHighChartModel();

                    graficoModel.name = item.DescMarca;
                    graficoModel.type = "line";


                    var listaDados = grafico.Where(a => a.DescMarca.Equals(item.DescMarca)).Select(a => new { a.Perc, a.Descricao, a.DescMarca, a.Base, a.Media, a.BaseMinima }).ToList();

                    foreach (var itemArray in listaDados)
                    {
                        graficoModel.data.Add(new DataModel
                        {
                            name = itemArray.DescMarca,
                            periodo = itemArray.Descricao,
                            valorbase = itemArray.Base,
                            y = itemArray.Perc,
                            media = itemArray.Media,
                            baseminima = itemArray.BaseMinima,

                        });
                    }

                    retorno.Grafico.Add(graficoModel);
                }
            }
            else
                return null;

            return retorno;
        }

        public static GraficoLinhasModel TratarDadosGraficoImagemEvolutivo(List<GraficoImagemEvolutivoProcedure> grafico)
        {
            var retorno = new GraficoLinhasModel();
            retorno.Cores = new List<string>();
            retorno.Periodos = new List<string>();
            retorno.Bases = new List<string>();
            retorno.Grafico = new List<GraficoLinhasHighChartModel>();

            if (grafico.Any())
            {

                var categorias = grafico.Select(a => new { a.DescMarca }).Distinct().ToList();


                foreach (var item in grafico.Select(a => new { a.DescOnda, a.BaseAbs }).Distinct().ToList())
                {

                    string descricao = string.Concat(item.DescOnda);
                    string valorBase = string.Concat(item.BaseAbs);
                    retorno.Bases.Add(valorBase);
                    retorno.Periodos.Add(descricao);
                }

                foreach (var item in categorias)
                {

                    var graficoModel = new GraficoLinhasHighChartModel();

                    graficoModel.name = item.DescMarca;
                    graficoModel.type = "line";


                    var listaDados = grafico.Where(a => a.DescMarca.Equals(item.DescMarca)).Select(a => new { a.Perc, a.DescOnda, a.DescMarca, a.BaseAbs, a.TesteSig, a.BaseMinima }).ToList();

                    foreach (var itemArray in listaDados)
                    {
                        graficoModel.data.Add(new DataModel
                        {
                            name = itemArray.DescMarca,
                            periodo = itemArray.DescOnda,
                            valorbase = itemArray.BaseAbs,
                            y = itemArray.Perc,
                            sig = itemArray.TesteSig,
                            baseminima = itemArray.BaseMinima,

                        });
                    }

                    retorno.Grafico.Add(graficoModel);
                }
            }
            else
                return null;

            return retorno;
        }


        /// <summary>
        /// EXEMPLO EXEMPLO EXEMPLO
        /// </summary>
        /// <param name="grafico"></param>
        /// <returns></returns>
        public static GraficoLinhasModel TratarDados(List<GraficoLinhasRetornoProcedurePadrao> grafico)
        {
            var retorno = new GraficoLinhasModel();
            retorno.Cores = new List<string>();
            retorno.Periodos = new List<string>();
            retorno.Bases = new List<string>();
            retorno.Grafico = new List<GraficoLinhasHighChartModel>();

            if (grafico.Any())
            {

                var categorias = grafico.Select(a => new { a.Descricao }).Distinct().ToList();


                foreach (var item in grafico.OrderBy(a => a.Descricao).Select(a => new { a.DescPeriodo, a.Base }).Distinct().ToList())
                {
                    string descricao = string.Concat(item.DescPeriodo);
                    string valorBase = string.Concat(item.Base);
                    retorno.Bases.Add(valorBase);
                    retorno.Periodos.Add(descricao);
                }

                foreach (var item in categorias)
                {

                    var graficoModel = new GraficoLinhasHighChartModel();

                    graficoModel.name = item.Descricao;
                    graficoModel.type = "line";


                    var listaDados = grafico.Where(a => a.Descricao.Equals(item.Descricao)).Select(a => new { a.Valor, a.DescPeriodo, a.Descricao, a.Base }).ToList();

                    foreach (var itemArray in listaDados)
                    {
                        graficoModel.data.Add(new DataModel
                        {
                            name = itemArray.Descricao,
                            periodo = itemArray.DescPeriodo,
                            valorbase = itemArray.Base,
                            y = itemArray.Valor,

                        });
                    }

                    retorno.Grafico.Add(graficoModel);
                }
            }
            else
                return null;

            return retorno;
        }




        public static TabelaPadraoAdHoc TratarDadosTabelaAdHocAtributo(List<TabelaAdHoc> dados, List<TabelaAdHocAtributo> titulos)
        {
            var retorno = new TabelaPadraoAdHoc();

            retorno.Titulos.AddRange(titulos);
            var bases = new TabelaAdHoc();

            if (dados.Any())
            {
                bases.Texto = "Base";
                bases.Perc1 = dados.FirstOrDefault().Base;
                bases.Perc2 = dados.FirstOrDefault().Base;
                bases.Perc3 = dados.FirstOrDefault().Base;
                bases.Perc4 = dados.FirstOrDefault().Base;
                bases.Perc5 = dados.FirstOrDefault().Base;
                bases.Perc6 = dados.FirstOrDefault().Base;
                bases.Perc7 = dados.FirstOrDefault().Base;
                bases.Perc8 = dados.FirstOrDefault().Base;
                bases.Perc9 = dados.FirstOrDefault().Base;
                bases.Perc10 = dados.FirstOrDefault().Base;

                retorno.Dados.Add(bases);
            }

            retorno.Dados.AddRange(dados);

            return retorno;

        }

    }
}