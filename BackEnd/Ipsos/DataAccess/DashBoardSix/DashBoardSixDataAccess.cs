using Business.TrataDados;
using Dapper;
using DataAccess.Config;
using Entities.GraficoBVCEvolutivo;
using Entities.GraficoBVCTop10;
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

namespace DataAccess.DashBoardSix
{
    public class DashBoardSixDataAccess
    {
        private readonly string usuarioEmail = string.Empty;

        public DashBoardSixDataAccess(string usuario)
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
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros1, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento1 = coluna;
                }

                var parametros2 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcel(filtro, filtro.Marca2, 2);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros2, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento2 = coluna;

                }

                var parametros3 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcel(filtro, filtro.Marca3, 3);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros3, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento3 = coluna;

                }

                var parametros4 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcel(filtro, filtro.Marca4, 4);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros4, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento4 = coluna;

                }

                var parametros5 = TrataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcel(filtro, filtro.Marca5, 5);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros5, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
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
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros1, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    if (filtro.Onda1 != null)
                        retorno.GraficoImagemPosicionamento1 = coluna;

                }

                var parametros2 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcel(filtro, filtro.Onda2, 7);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros2, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    if (filtro.Onda2 != null)
                        retorno.GraficoImagemPosicionamento2 = coluna;

                }

                var parametros3 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcel(filtro, filtro.Onda3, 8);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros3, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    if (filtro.Onda3 != null)
                        retorno.GraficoImagemPosicionamento3 = coluna;

                }

                var parametros4 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcel(filtro, filtro.Onda4, 9);
                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros4, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    if (filtro.Onda4 != null)
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
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento1 = coluna;
                }

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna1_2, filtro.MarcaDuploColuna1, 8);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento2 = coluna;
                }

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna2, filtro.MarcaDuploColuna2, 9);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento3 = coluna;
                }

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna2_2, filtro.MarcaDuploColuna2, 10);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento4 = coluna;
                }

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna3, filtro.MarcaDuploColuna3, 11);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento5 = coluna;
                }

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna3_2, filtro.MarcaDuploColuna3, 12);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento6 = coluna;
                }

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna4, filtro.MarcaDuploColuna4, 13);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento7 = coluna;
                }

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna4_2, filtro.MarcaDuploColuna4, 14);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento8 = coluna;
                }

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna5, filtro.MarcaDuploColuna5, 15);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento9 = coluna;
                }

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var parametros = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(filtro, filtro.OndaDuploColuna5_2, filtro.MarcaDuploColuna5, 16);
                    var coluna = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                    retorno.GraficoImagemPosicionamento10 = coluna;
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

                if (filtro.Onda[0] == null)
                {
                    filtro.Onda = null;
                }

                var TrataFiltros = new TrataFiltros();
                //var parametros = TrataFiltros.MontaParametrosFiltroPadraoEEfeitos(filtro);

                var parametros = new DynamicParameters();
                parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null) ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
                parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
                parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
                parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
                parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));
                parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
                parametros.Add("@ParamListaMarca", filtro.Marca == null ? "" : string.Join(",", filtro.Marca.Select(s => string.Concat(s.IdItem))));

                parametros.Add("@ParamCodUser", filtro.CodUser);
                parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
                parametros.Add("@ParamSequencia", filtro.Sequencia);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    retorno = conexaoBD.Query<GraficoImagemPosicionamento>("pr_Dashboard_BVCMercado", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    if (retorno.Any())
                    {
                        var teste = retorno;
                    }
                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }


        public List<GraficoBVCTop10> CarregarGraficoBVCTop10Marcas(FiltroPadrao filtro)
        {
            var retorno = new List<GraficoBVCTop10>();

            try
            {

                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadrao(filtro);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    retorno = conexaoBD.Query<GraficoBVCTop10>("pr_Dashboard_BVCTop10", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }


        public List<GraficoBVCEvolutivo> CarregarGraficoBVCEvolutivo(FiltroPadrao filtro)
        {
            var retorno = new List<GraficoBVCEvolutivo>();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroBVCEvolutivo(filtro);

                if (filtro.Onda == null)
                    return retorno;

                parametros.Add("@ParamOndaAtual", filtro.ParamOndaAtual);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    retorno = conexaoBD.Query<GraficoBVCEvolutivo>("pr_Dashboard_BVCEvolutivo", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }


        public GraficoBVCEvolutivoFull CarregarGraficoBVCEvolutivoExcel(FiltroPadraoExcel filtro)
        {
            var retorno = new GraficoBVCEvolutivoFull();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros1 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcelTop(filtro, filtro.Onda1, 1);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    if (filtro.Onda1[0] != null)
                    retorno.GraficoBVCEvolutivo1 = conexaoBD.Query<GraficoBVCEvolutivo>("pr_Dashboard_BVCEvolutivo", parametros1, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                }

                var parametros2 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcelTop(filtro, filtro.Onda2, 2);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    if (filtro.Onda2[0] != null)
                        retorno.GraficoBVCEvolutivo2 = conexaoBD.Query<GraficoBVCEvolutivo>("pr_Dashboard_BVCEvolutivo", parametros2, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                }

                var parametros3 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcelTop(filtro, filtro.Onda3, 3);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    if (filtro.Onda3[0] != null)
                        retorno.GraficoBVCEvolutivo3 = conexaoBD.Query<GraficoBVCEvolutivo>("pr_Dashboard_BVCEvolutivo", parametros3, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
                }

                var parametros4 = TrataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcelTop(filtro, filtro.Onda4, 4);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    if (filtro.Onda4[0] != null)
                        retorno.GraficoBVCEvolutivo4 = conexaoBD.Query<GraficoBVCEvolutivo>("pr_Dashboard_BVCEvolutivo", parametros4, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();
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
