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

namespace DataAccess.DashBoardTwo
{
    public class DashBoardTwoDataAccess
    {
        private readonly string usuarioEmail = string.Empty;

        public DashBoardTwoDataAccess(string usuario)
        {
            usuarioEmail = usuario;
        }

        #region  MÉTODOS PARA GERAÇÃO DE DADOS PARA EXCEL
        public GraficoColunasFullLoad CarregarGraficoComparativoMarcasExcel(FiltroPadraoExcel filtro)
        {
            var retorno = new GraficoColunasFullLoad();

            try
            {

                var TrataFiltros = new TrataFiltros();
                var parametros1 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca1,1);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Awareness", parametros1, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoColunas1 = coluna.FirstOrDefault();

                }

                var parametros2 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca2,2);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Awareness", parametros2, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoColunas2 = coluna.FirstOrDefault();

                }

                var parametros3 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca3,3);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Awareness", parametros3, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoColunas3 = coluna.FirstOrDefault();

                }

                var parametros4 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca4,4);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Awareness", parametros4, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoColunas4 = coluna.FirstOrDefault();

                }

                var parametros5 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca5,5);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Awareness", parametros5, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

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

        public GraficoColunasFullLoad CarregarGraficoEvolutivoMarcasExcel(FiltroPadraoExcel filtro)
        {
            var retorno = new GraficoColunasFullLoad();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros1 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcelDenominator(filtro, filtro.Onda1,6);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Awareness", parametros1, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoColunas1 = coluna.FirstOrDefault();

                }

                var parametros2 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcelDenominator(filtro, filtro.Onda2,7);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Awareness", parametros2, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoColunas2 = coluna.FirstOrDefault();

                }

                var parametros3 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcelDenominator(filtro, filtro.Onda3,8);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Awareness", parametros3, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoColunas3 = coluna.FirstOrDefault();

                }

                var parametros4 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcelDenominator(filtro, filtro.Onda4,9);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoColunas>("pr_Dashboard_Awareness", parametros4, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

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


        public GraficoColunas CarregarGraficoComparativoMarcas(FiltroPadrao filtro)
        {
            var retorno = new GraficoColunas();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoDenominator(filtro);


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
