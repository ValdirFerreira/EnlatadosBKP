using Business.TrataDados;
using Dapper;
using DataAccess.Config;
using Entities.Filtros;
using Entities.GraficoAranha;
using Entities.GraficoHighChart;
using Entities.GraficoImagemEvolutivo;
using Entities.GraficoLinhasImagem;
using Entities.Parametros;
using Entities.TabelaAdHoc;
using Helpers.Logtxt;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess.DashBoardNine
{
    public class DashBoardNineDataAccess
    {
        private readonly string usuarioEmail = string.Empty;

        public DashBoardNineDataAccess(string usuario)
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
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoNine(filtro);


                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var listRetornoProcedure = conexaoBD.Query<GraficoLinhasImagemProcedure>("pr_Dashboard_AdHoc_Bloco2", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();


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
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoImagemAdHocCustom(filtro);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var listRetornoProcedure = conexaoBD.Query<GraficoImagemEvolutivoProcedure>("pr_Dashboard_AdHoc_Bloco1", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();


                    if (listRetornoProcedure.Any())
                    {
                        retorno = TratamentoDados.TratarDadosGraficoImagemEvolutivo(listRetornoProcedure);

                       // retorno.Cores.AddRange(listRetornoProcedure.Select(a => a.CorSiteMarca).Distinct().ToList());
                    }

                    retorno.Cores.Add("#4056B7");
                    retorno.Cores.Add("#F2C94C");
                    retorno.Cores.Add("#EB5757");

                }


            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }

        public GraficoLinhasModel ImagemEvolutivaLinhas2(FiltroPadrao filtro)
        {
            var retorno = new GraficoLinhasModel();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoImagemAdHocSemMarca(filtro);
                parametros.Add("@ParamCodMarca", filtro.Marca.FirstOrDefault().IdItem);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var listRetornoProcedure = conexaoBD.Query<GraficoImagemEvolutivoProcedure>("pr_Dashboard_AdHoc_Bloco11", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();


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



        public TabelaPadraoAdHoc TabelaAdHocAtributo(FiltroPadrao filtro)
        {
            var retorno = new TabelaPadraoAdHoc();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoNine(filtro);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
      
                    var tables = conexaoBD.QueryMultipleAsync("pr_Dashboard_AdHoc_Bloco4 ", parametros, null, 300, System.Data.CommandType.StoredProcedure).Result;

                    var table1 = tables.Read<TabelaAdHoc>().ToList();
                    var table2 = tables.Read<TabelaAdHocAtributo>().ToList();

                    retorno = TratamentoDados.TratarDadosTabelaAdHocAtributo(table1, table2);

                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }

        public TabelaPadraoAdHoc TabelaAdHocAtributoBloco6(FiltroPadrao filtro)
        {
            var retorno = new TabelaPadraoAdHoc();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoNine(filtro);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {

                    var tables = conexaoBD.QueryMultipleAsync("pr_Dashboard_AdHoc_Bloco6 ", parametros, null, 300, System.Data.CommandType.StoredProcedure).Result;

                    var table1 = tables.Read<TabelaAdHoc>().ToList();
                    var table2 = tables.Read<TabelaAdHocAtributo>().ToList();

                    retorno = TratamentoDados.TratarDadosTabelaAdHocAtributo(table1, table2);

                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }



        public TabelaPadraoAdHoc TabelaAdHocAtributoBloco10(FiltroPadrao filtro)
        {
            var retorno = new TabelaPadraoAdHoc();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoNine(filtro);

                parametros.Add("@ParamCodMarca", filtro.Marca.FirstOrDefault().IdItem);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {

                    var tables = conexaoBD.QueryMultipleAsync("pr_Dashboard_AdHoc_Bloco10", parametros, null, 300, System.Data.CommandType.StoredProcedure).Result;

                    var table1 = tables.Read<TabelaAdHoc>().ToList();
                    var table2 = tables.Read<TabelaAdHocAtributo>().ToList();

                    retorno = TratamentoDados.TratarDadosTabelaAdHocAtributo(table1, table2);

                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }


        public TabelaPadraoAdHoc TabelaAdHocAtributoBloco2(FiltroPadrao filtro)
        {
            var retorno = new TabelaPadraoAdHoc();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoNine(filtro);


                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {

                    var tables = conexaoBD.QueryMultipleAsync("pr_Dashboard_AdHoc_Bloco2", parametros, null, 300, System.Data.CommandType.StoredProcedure).Result;

                    var table1 = tables.Read<TabelaAdHoc>().ToList();
                    var table2 = tables.Read<TabelaAdHocAtributo>().ToList();

                    retorno = TratamentoDados.TratarDadosTabelaAdHocAtributo(table1, table2);

                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }


        public GraficoLinhasImagem CarregarGraficoImagemEvolutivaNineExcel(FiltroPadraoExcel filtro)
        {
            var retorno = new GraficoLinhasImagem();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros1 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelNine(filtro, filtro.Marca1, 1);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoLinhasImagemProcedure>("pr_Dashboard_AdHoc_Bloco2", parametros1, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoLinhasImagem1 = coluna;

                }

                var parametros2 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelNine(filtro, filtro.Marca2, 2);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoLinhasImagemProcedure>("pr_Dashboard_AdHoc_Bloco2", parametros2, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoLinhasImagem2 = coluna;

                }

                var parametros3 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelNine(filtro, filtro.Marca3, 3);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoLinhasImagemProcedure>("pr_Dashboard_AdHoc_Bloco2", parametros3, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoLinhasImagem3 = coluna;

                }

                var parametros4 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelNine(filtro, filtro.Marca4, 4);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoLinhasImagemProcedure>("pr_Dashboard_AdHoc_Bloco2", parametros4, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoLinhasImagem4 = coluna;

                }

                var parametros5 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelNine(filtro, filtro.Marca5, 5);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoLinhasImagemProcedure>("pr_Dashboard_AdHoc_Bloco2", parametros5, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

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

        public List<GraficoImagemEvolutivoProcedure> ImagemEvolutivaLinhasNineExcel(FiltroPadrao filtro)
        {
            var retorno = new List<GraficoImagemEvolutivoProcedure>();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoImagemAdHoc(filtro);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var listRetornoProcedure = conexaoBD.Query<GraficoImagemEvolutivoProcedure>("pr_Dashboard_AdHoc_Bloco1", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    return listRetornoProcedure;

                }


            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }

        public List<GraficoImagemEvolutivoProcedure> ImagemEvolutivaLinhasNineExcel2(FiltroPadrao filtro)
        {
            var retorno = new List<GraficoImagemEvolutivoProcedure>();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoImagemAdHocSemMarca(filtro);
                parametros.Add("@ParamCodMarca", filtro.Marca.FirstOrDefault().IdItem);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var listRetornoProcedure = conexaoBD.Query<GraficoImagemEvolutivoProcedure>("pr_Dashboard_AdHoc_Bloco11", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    return listRetornoProcedure;

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
