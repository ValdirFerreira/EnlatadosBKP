using Business.TrataDados;
using Dapper;
using DataAccess.Config;
using Entities.Filtros;
using Entities.GraficoAranha;
using Entities.GraficoHighChart;
using Entities.GraficoImagemEvolutivo;
using Entities.GraficoLinhasImagem;
using Entities.Parametros;
using Helpers.Logtxt;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess.DashBoardFour
{
    public class DashBoardFourDataAccess
    {
        private readonly string usuarioEmail = string.Empty;

        public DashBoardFourDataAccess(string usuario)
        {
            usuarioEmail = usuario;
        }



        public List<GraficoAranhaRetornoProcedure> ComparativoMarcasExcel(FiltroPadrao filtro)
        {
            var retorno = new List<GraficoAranhaRetornoProcedure>();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoImagem(filtro);


                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var listRetornoProcedure = conexaoBD.Query<GraficoAranhaRetornoProcedure>("pr_Dashboard_ImagemMarcas", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    return listRetornoProcedure;

                }
            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }
     
        public GraficoLinhasImagem CarregarGraficoImagemEvolutivaExcel(FiltroPadraoExcel filtro)
        {
            var retorno = new GraficoLinhasImagem();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros1 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelBIA(filtro, filtro.Marca1, 1);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoLinhasImagemProcedure>("pr_Dashboard_Imagem2Periodos", parametros1, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoLinhasImagem1 = coluna;

                }

                var parametros2 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelBIA(filtro, filtro.Marca2, 2);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoLinhasImagemProcedure>("pr_Dashboard_Imagem2Periodos", parametros2, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoLinhasImagem2 = coluna;

                }

                var parametros3 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelBIA(filtro, filtro.Marca3, 3);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoLinhasImagemProcedure>("pr_Dashboard_Imagem2Periodos", parametros3, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoLinhasImagem3 = coluna;

                }

                var parametros4 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelBIA(filtro, filtro.Marca4, 4);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoLinhasImagemProcedure>("pr_Dashboard_Imagem2Periodos", parametros4, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoLinhasImagem4 = coluna;

                }

                var parametros5 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelBIA(filtro, filtro.Marca5, 5);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoLinhasImagemProcedure>("pr_Dashboard_Imagem2Periodos", parametros5, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoLinhasImagem5 = coluna;

                }

               

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }



        public List<GraficoImagemEvolutivoProcedure> ImagemEvolutivaLinhasExcel(FiltroPadrao filtro)
        {
            var retorno = new List<GraficoImagemEvolutivoProcedure>();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoImagem(filtro);
                parametros.Add("@ParamAtributo", filtro.Atributo);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var listRetornoProcedure = conexaoBD.Query<GraficoImagemEvolutivoProcedure>("pr_Dashboard_ImagemEvolutivo", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    return listRetornoProcedure;

                }


            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }


        public GraficoLinhasModel ComparativoMarcas(FiltroPadrao filtro)
        {
            var retorno = new GraficoLinhasModel();

            try
            {

                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoImagem(filtro);


                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var listRetornoProcedure = conexaoBD.Query<GraficoAranhaRetornoProcedure>("pr_Dashboard_ImagemMarcas", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    retorno = TratamentoDados.TratarDadosGraficoAranha(listRetornoProcedure);

                    retorno.Cores.AddRange(listRetornoProcedure.Select(a => a.CorSiteMarca).Distinct().ToList());

                    //retorno.Cores.Add("#4056B7");
                    //retorno.Cores.Add("#F2C94C");
                    //retorno.Cores.Add("#EB5757");

                }
            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }

        public GraficoLinhasModel ImagemEvolutiva(FiltroPadrao filtro)
        {
            var retorno = new GraficoLinhasModel();

            try
            {

                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoImagem(filtro);


                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var listRetornoProcedure = conexaoBD.Query<GraficoLinhasImagemProcedure>("pr_Dashboard_Imagem2Periodos", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();


                    if (listRetornoProcedure.Any())
                    {
                        retorno = TratamentoDados.TratarDadosGraficoLinhasImagem(listRetornoProcedure);

                        retorno.Cores.Add("#C4C4C4");
                        retorno.Cores.Add(listRetornoProcedure.FirstOrDefault().CorSiteMarca);
                    }

                    //retorno.Cores.Add("#4056B7");
                    //retorno.Cores.Add("#F2C94C");
                    //retorno.Cores.Add("#EB5757");

                }
            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }

        public GraficoLinhasModel ImagemEvolutivaLinhas(FiltroPadrao filtro)
        {
            var retorno = new GraficoLinhasModel();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoImagem(filtro);
                parametros.Add("@ParamAtributo", filtro.Atributo);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var listRetornoProcedure = conexaoBD.Query<GraficoImagemEvolutivoProcedure>("pr_Dashboard_ImagemEvolutivo", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();


                    if (listRetornoProcedure.Any())
                    {
                        retorno = TratamentoDados.TratarDadosGraficoImagemEvolutivo(listRetornoProcedure);

                        retorno.Cores.AddRange(listRetornoProcedure.Select(a => a.CorSiteMarca).Distinct().ToList());
                    }

                    //retorno.Cores.Add("#4056B7");
                    //retorno.Cores.Add("#F2C94C");
                    //retorno.Cores.Add("#EB5757");

                }


            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }
    }
}
