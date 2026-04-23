using Business.TrataDados;
using Dapper;
using DataAccess.Config;
using Entities.GraficoFunil;
using Entities.Parametros;
using Helpers.Logtxt;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess.DashboardThree
{
    public class DashboardThreeDataAccess
    {

        private readonly string usuarioEmail = string.Empty;

        public DashboardThreeDataAccess(string usuario)
        {
            usuarioEmail = usuario;
        }


        #region  MÉTODOS PARA GERAÇÃO DE DADOS PARA EXCEL
        public GraficoFunilFullLoad CarregarGraficoComparativoMarcasExcel(FiltroPadraoExcel filtro)
        {
            var retorno = new GraficoFunilFullLoad();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros1 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca1, 1);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoFunil>("pr_Dashboard_Funil", parametros1, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoFunil1 = coluna.FirstOrDefault();

                }

                var parametros2 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca2, 2);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoFunil>("pr_Dashboard_Funil", parametros2, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoFunil2 = coluna.FirstOrDefault();

                }

                var parametros3 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca3, 3);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoFunil>("pr_Dashboard_Funil", parametros3, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoFunil3 = coluna.FirstOrDefault();

                }

                var parametros4 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca4, 4);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoFunil>("pr_Dashboard_Funil", parametros4, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoFunil4 = coluna.FirstOrDefault();

                }

                var parametros5 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca5, 5);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoFunil>("pr_Dashboard_Funil", parametros5, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoFunil5 = coluna.FirstOrDefault();

                }

                var parametros6 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca6, 6);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoFunil>("pr_Dashboard_Funil", parametros6, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoFunil6 = coluna.FirstOrDefault();

                }

                var parametros7 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca7, 7);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoFunil>("pr_Dashboard_Funil", parametros7, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoFunil7 = coluna.FirstOrDefault();

                }

                var parametros8 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca8, 8);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoFunil>("pr_Dashboard_Funil", parametros8, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (coluna.Count > 0)
                        retorno.GraficoFunil8 = coluna.FirstOrDefault();

                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }

        public GraficoFunilFullLoad CarregarGraficoEvolutivoMarcasExcel(FiltroPadraoExcel filtro)
        {
            var retorno = new GraficoFunilFullLoad();

            try
            {
                var TrataFiltros = new TrataFiltros();

                var parametros1 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcelDenominator(filtro, filtro.Onda1, 6);


                if (filtro.Onda1 != null && filtro.Onda1[0] != null)
                    using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                    {
                        var coluna = conexaoBD.Query<GraficoFunil>("pr_Dashboard_Funil", parametros1, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                        if (coluna.Count > 0)
                            retorno.GraficoFunil1 = coluna.FirstOrDefault();

                    }

                var parametros2 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcelDenominator(filtro, filtro.Onda2, 7);
                if (filtro.Onda2 != null && filtro.Onda2[0] != null)
                    using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                    {
                        var coluna = conexaoBD.Query<GraficoFunil>("pr_Dashboard_Funil", parametros2, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                        if (coluna.Count > 0)
                            retorno.GraficoFunil2 = coluna.FirstOrDefault();

                    }

                var parametros3 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcelDenominator(filtro, filtro.Onda3, 8);
                if (filtro.Onda3 != null && filtro.Onda3[0] != null)
                    using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                    {
                        var coluna = conexaoBD.Query<GraficoFunil>("pr_Dashboard_Funil", parametros3, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                        if (coluna.Count > 0)
                            retorno.GraficoFunil3 = coluna.FirstOrDefault();

                    }

                var parametros4 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcelDenominator(filtro, filtro.Onda4, 9);
                if (filtro.Onda4 != null && filtro.Onda4[0] != null)
                    using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                    {
                        var coluna = conexaoBD.Query<GraficoFunil>("pr_Dashboard_Funil", parametros4, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                        if (coluna.Count > 0)
                            retorno.GraficoFunil4 = coluna.FirstOrDefault();

                    }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }

        #endregion


        public GraficoFunil CarregarGraficoComparativoExperiencia(FiltroPadrao filtro)
        {
            var retorno = new GraficoFunil();

            try
            {

                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoDenominator(filtro);


                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var list = conexaoBD.Query<GraficoFunil>("pr_Dashboard_Funil", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

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
