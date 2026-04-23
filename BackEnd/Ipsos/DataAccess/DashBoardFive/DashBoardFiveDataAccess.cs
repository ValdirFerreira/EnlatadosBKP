using Business.TrataDados;
using Dapper;
using DataAccess.Config;
using Entities.GraficoColunas;
using Entities.GraficoImagemPosicionamento;
using Entities.Parametros;
using Entities.TraducaoIdioma;
using Helpers.Logtxt;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess.DashBoardFive
{
    public class DashBoardFiveDataAccess
    {
        private readonly string usuarioEmail = string.Empty;

        public DashBoardFiveDataAccess(string usuario)
        {
            usuarioEmail = usuario;
        }

        #region  MÉTODOS PARA GERAÇÃO DE DADOS PARA EXCEL
        public GraficoImagemPosicionamentoFullLoad CarregarGraficoComparativoMarcasExcel(FiltroPadraoExcel filtro)
        {
            var retorno = new GraficoImagemPosicionamentoFullLoad();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros1 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcel(filtro, filtro.Marca1, 1);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamentoDinamico", parametros1, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento1 = coluna;
                }

                var parametros2 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcel(filtro, filtro.Marca2, 2);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamentoDinamico", parametros2, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento2 = coluna;

                }

                var parametros3 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcel(filtro, filtro.Marca3, 3);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamentoDinamico", parametros3, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento3 = coluna;

                }

                var parametros4 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcel(filtro, filtro.Marca4, 4);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamentoDinamico", parametros4, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento4 = coluna;

                }

                var parametros5 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcel(filtro, filtro.Marca5, 5);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamentoDinamico", parametros5, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento5 = coluna;

                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }

        public GraficoImagemPosicionamentoFullLoad CarregarGraficoEvolutivoMarcasExcel(FiltroPadraoExcel filtro)
        {
            var retorno = new GraficoImagemPosicionamentoFullLoad();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros1 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcel(filtro, filtro.Onda1, 6);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros1, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento1 = coluna;

                }

                var parametros2 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcel(filtro, filtro.Onda2, 7);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros2, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento2 = coluna;

                }

                var parametros3 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcel(filtro, filtro.Onda3, 8);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros3, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento3 = coluna;

                }

                var parametros4 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcel(filtro, filtro.Onda4, 9);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros4, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento4 = coluna;

                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }


        public GraficoImagemPosicionamentoFullLoad CarregarGraficoEvolutivoMarcasDuploExcel(FiltroPadraoExcelDuplo filtro)
        {
            var retorno = new GraficoImagemPosicionamentoFullLoad();

            try
            {
                var TrataFiltros = new TrataFiltros();

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna1, filtro.MarcaDuploColuna1, 7);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento1 = coluna;
                }

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna1_2, filtro.MarcaDuploColuna1, 8);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento2 = coluna;
                }

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna1_3, filtro.MarcaDuploColuna1, 7);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento1_3 = coluna;
                }



                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna2, filtro.MarcaDuploColuna2, 9);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento3 = coluna;
                }

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna2_2, filtro.MarcaDuploColuna2, 10);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento4 = coluna;
                }
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna2_3, filtro.MarcaDuploColuna2, 9);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento3_3 = coluna;
                }


                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna3, filtro.MarcaDuploColuna3, 11);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento5 = coluna;
                }

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna3_2, filtro.MarcaDuploColuna3, 12);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento6 = coluna;
                }
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna3_3, filtro.MarcaDuploColuna3, 11);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento5_3 = coluna;
                }


                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna4, filtro.MarcaDuploColuna4, 13);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento7 = coluna;
                }

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna4_2, filtro.MarcaDuploColuna4, 14);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento8 = coluna;
                }
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna4_3, filtro.MarcaDuploColuna4, 13);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento7_3 = coluna;
                }


                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna5, filtro.MarcaDuploColuna5, 15);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento9 = coluna;
                }

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna5_2, filtro.MarcaDuploColuna5, 16);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento10 = coluna;
                }
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna5_3, filtro.MarcaDuploColuna5, 15);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento9_3 = coluna;
                }


            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }

        #endregion


        public List<GraficoImagemPosicionamento> CarregarGraficoComparativoMarcas(FiltroPadrao filtro)
        {
            var retorno = new List<GraficoImagemPosicionamento>();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadrao(filtro);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    //retorno = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamentoDinamico", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }

        public List<GraficoImagemPosicionamento> CarregarGraficoComparativoMarcas2(FiltroPadrao filtro)
        {
            var retorno = new List<GraficoImagemPosicionamento>();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadrao(filtro);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    retorno = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_ImagemPosicionamento", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                }


                foreach (var item in retorno)
                {
                    if(!string.IsNullOrEmpty( item.BaseMinima))
                    {
                        var teste = " ";
                    }
                }
            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }

        public GraficoColunas CarregarGraficoEvolutivoMarcas(FiltroPadrao filtro)
        {
            var retorno = new GraficoColunas();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadrao(filtro);


                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var list = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Awareness", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    var trataDados = new TrataDadosDashBoardTwo();

                    if (list.Count > 0)
                        retorno = list.FirstOrDefault();

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
