using Business.TrataDados;
using Dapper;
using DataAccess.Config;

using Entities.GraficoColunas;
using Entities.Parametros;
using Helpers.Logtxt;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;

namespace DataAccess.DashBoardTwo
{
    public class ComunicacaoLikeDislikeDataAccess
    {
        private readonly string usuarioEmail = string.Empty;

        public ComunicacaoLikeDislikeDataAccess(string usuario)
        {
            usuarioEmail = usuario;
        }

        #region MÉTODOS PARA GERAÇÃO DE DADOS PARA EXCEL

        public ComunicacaoLikeDislikeFullLoad CarregarComparativoMarcasExcel(FiltroPadraoExcel filtro)
        {
            var retorno = new ComunicacaoLikeDislikeFullLoad();

            try
            {
                var trataFiltros = new TrataFiltros();

                // Marca 1
                var parametros1 = trataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca1, 1);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var dados = conexaoBD.Query<ComunicacaoLikeDislike>(
                        "pr_Dashboard_ComunicacaoLikeDislike",
                        parametros1,
                        null,
                        false,
                        300,
                        System.Data.CommandType.StoredProcedure).ToList();

                    if (dados.Count > 0)
                        retorno.ComunicacaoLikeDislike1 = dados.FirstOrDefault();
                }

                // Marca 2
                var parametros2 = trataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca2, 2);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var dados = conexaoBD.Query<ComunicacaoLikeDislike>(
                        "pr_Dashboard_ComunicacaoLikeDislike",
                        parametros2,
                        null,
                        false,
                        300,
                        System.Data.CommandType.StoredProcedure).ToList();

                    if (dados.Count > 0)
                        retorno.ComunicacaoLikeDislike2 = dados.FirstOrDefault();
                }

                // Marca 3
                var parametros3 = trataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca3, 3);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var dados = conexaoBD.Query<ComunicacaoLikeDislike>(
                        "pr_Dashboard_ComunicacaoLikeDislike",
                        parametros3,
                        null,
                        false,
                        300,
                        System.Data.CommandType.StoredProcedure).ToList();

                    if (dados.Count > 0)
                        retorno.ComunicacaoLikeDislike3 = dados.FirstOrDefault();
                }

                // Marca 4
                var parametros4 = trataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca4, 4);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var dados = conexaoBD.Query<ComunicacaoLikeDislike>(
                        "pr_Dashboard_ComunicacaoLikeDislike",
                        parametros4,
                        null,
                        false,
                        300,
                        System.Data.CommandType.StoredProcedure).ToList();

                    if (dados.Count > 0)
                        retorno.ComunicacaoLikeDislike4 = dados.FirstOrDefault();
                }

                // Marca 5
                var parametros5 = trataFiltros.MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(filtro, filtro.Marca5, 5);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var dados = conexaoBD.Query<ComunicacaoLikeDislike>(
                        "pr_Dashboard_ComunicacaoLikeDislike",
                        parametros5,
                        null,
                        false,
                        300,
                        System.Data.CommandType.StoredProcedure).ToList();

                    if (dados.Count > 0)
                        retorno.ComunicacaoLikeDislike5 = dados.FirstOrDefault();
                }
            }
            catch (Exception ex)
            {
                LogText.Instance.Error(
                    this.GetType().Name,
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }

        public ComunicacaoLikeDislikeFullLoad CarregarEvolutivoMarcasExcel(FiltroPadraoExcel filtro)
        {
            var retorno = new ComunicacaoLikeDislikeFullLoad();

            try
            {
                var trataFiltros = new TrataFiltros();

                // Onda 1
                var parametros1 = trataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcelDenominator(filtro, filtro.Onda1, 6);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var dados = conexaoBD.Query<ComunicacaoLikeDislike>(
                        "pr_Dashboard_ComunicacaoLikeDislike",
                        parametros1,
                        null,
                        false,
                        300,
                        System.Data.CommandType.StoredProcedure).ToList();

                    if (dados.Count > 0)
                        retorno.ComunicacaoLikeDislike1 = dados.FirstOrDefault();
                }

                // Onda 2
                var parametros2 = trataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcelDenominator(filtro, filtro.Onda2, 7);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var dados = conexaoBD.Query<ComunicacaoLikeDislike>(
                        "pr_Dashboard_ComunicacaoLikeDislike",
                        parametros2,
                        null,
                        false,
                        300,
                        System.Data.CommandType.StoredProcedure).ToList();

                    if (dados.Count > 0)
                        retorno.ComunicacaoLikeDislike2 = dados.FirstOrDefault();
                }

                // Onda 3
                var parametros3 = trataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcelDenominator(filtro, filtro.Onda3, 8);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var dados = conexaoBD.Query<ComunicacaoLikeDislike>(
                        "pr_Dashboard_ComunicacaoLikeDislike",
                        parametros3,
                        null,
                        false,
                        300,
                        System.Data.CommandType.StoredProcedure).ToList();

                    if (dados.Count > 0)
                        retorno.ComunicacaoLikeDislike3 = dados.FirstOrDefault();
                }

                // Onda 4
                var parametros4 = trataFiltros.MontaParametrosFiltroPadraoEvolutivoMarcasExcelDenominator(filtro, filtro.Onda4, 9);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var dados = conexaoBD.Query<ComunicacaoLikeDislike>(
                        "pr_Dashboard_ComunicacaoLikeDislike",
                        parametros4,
                        null,
                        false,
                        300,
                        System.Data.CommandType.StoredProcedure).ToList();

                    if (dados.Count > 0)
                        retorno.ComunicacaoLikeDislike4 = dados.FirstOrDefault();
                }
            }
            catch (Exception ex)
            {
                LogText.Instance.Error(
                    this.GetType().Name,
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }

        #endregion

        public ComunicacaoLikeDislike CarregarComparativoMarcas(FiltroPadrao filtro)
        {
            var retorno = new ComunicacaoLikeDislike();

            try
            {

                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoComunicacao(filtro);
                parametros.Add("@ParamSTB", filtro.ParamSTB);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var list = conexaoBD.Query<ComunicacaoLikeDislike>(
                        "pr_Dashboard_ComunicacaoLikeDislike",
                        parametros,
                        null,
                        false,
                        300,
                        System.Data.CommandType.StoredProcedure).ToList();

                    if (list.Count > 0)
                        retorno = list.FirstOrDefault();
                }
            }
            catch (Exception ex)
            {
                LogText.Instance.Error(
                    this.GetType().Name,
                    System.Reflection.MethodBase.GetCurrentMethod().Name,
                    "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }
    }
}