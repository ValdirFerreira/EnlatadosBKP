using Business.TrataDados;
using Dapper;
using DataAccess.Config;
using Entities.GraficoColunas;
using Entities.Parametros;
using Entities.TraducaoIdioma;
using Helpers.Logtxt;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess.DashBoardEight
{
    public class DashBoardEightDataAccess
    {
        private readonly string usuarioEmail = string.Empty;

        public DashBoardEightDataAccess(string usuario)
        {
            usuarioEmail = usuario;
        }

        #region  MÉTODOS PARA GERAÇÃO DE DADOS PARA EXCEL
        public GraficoColunasFullLoad CarregarGraficoComparativoMarcasConsiderationExcel(FiltroPadraoExcel filtro)
        {
            var retorno = new GraficoColunasFullLoad();

            try
            {

                var TrataFiltros = new TrataFiltros();
                var parametros1 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcel(filtro, filtro.Marca1,1);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Consideration", parametros1, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoColunas1 = coluna.FirstOrDefault();

                }

                var parametros2 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcel(filtro, filtro.Marca2,2);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Consideration", parametros2, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoColunas2 = coluna.FirstOrDefault();

                }

                var parametros3 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcel(filtro, filtro.Marca3,3);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Consideration", parametros3, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoColunas3 = coluna.FirstOrDefault();

                }

                var parametros4 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcel(filtro, filtro.Marca4,4);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Consideration", parametros4, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoColunas4 = coluna.FirstOrDefault();

                }

                var parametros5 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcel(filtro, filtro.Marca5,5);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Consideration", parametros5, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoColunas5 = coluna.FirstOrDefault();

                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }

        public GraficoColunasFullLoad CarregarGraficoEvolutivoMarcasConsiderationExcel(FiltroPadraoExcel filtro)
        {
            var retorno = new GraficoColunasFullLoad();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros1 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcel(filtro, filtro.Onda1,6);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Consideration", parametros1, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoColunas1 = coluna.FirstOrDefault();

                }

                var parametros2 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcel(filtro, filtro.Onda2,7);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Consideration", parametros2, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoColunas2 = coluna.FirstOrDefault();

                }

                var parametros3 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcel(filtro, filtro.Onda3,8);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Consideration", parametros3, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoColunas3 = coluna.FirstOrDefault();

                }

                var parametros4 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcel(filtro, filtro.Onda4,9);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Consideration", parametros4, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoColunas4 = coluna.FirstOrDefault();

                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }

        #endregion


        public GraficoColunas CarregarGraficoComparativoMarcasConsideration(FiltroPadrao filtro)
        {
            var retorno = new GraficoColunas();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadrao(filtro);


                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var list = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Consideration", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

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
