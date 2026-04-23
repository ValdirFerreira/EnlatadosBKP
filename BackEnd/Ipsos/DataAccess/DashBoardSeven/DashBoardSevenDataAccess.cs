using Business.TrataDados;
using Dapper;
using DataAccess.Config;
using Entities.ComunicacaoQuadroResumo;
using Entities.GraficoColunas;
using Entities.GraficoComunicacaoDiagnostico;
using Entities.GraficoComunicacaoRecall;
using Entities.GraficoComunicacaoVisto;
using Entities.Parametros;
using Entities.TraducaoIdioma;
using Helpers.Logtxt;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess.DashBoardSeven
{
    public class DashBoardSevenDataAccess
    {
        private readonly string usuarioEmail = string.Empty;

        public DashBoardSevenDataAccess(string usuario)
        {
            usuarioEmail = usuario;
        }


        public List<GraficoComunicacaoRecall> CarregarGraficoComunicacaoRecall(FiltroPadrao filtro)
        {
            var retorno = new List<GraficoComunicacaoRecall>();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoComunicacao(filtro);
                parametros.Add("@ParamSTB", filtro.ParamSTB);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var list = conexaoBD.Query<GraficoComunicacaoRecall>("pr_Dashboard_ComunicacaoRecall", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    retorno = list;
                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }

        public List<GraficoComunicacaoVisto> CarregarGraficoComunicacaoVisto(FiltroPadrao filtro)
        {
            var retorno = new List<GraficoComunicacaoVisto>();

            try
            {

                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoComunicacao(filtro);
                parametros.Add("@ParamSTB", filtro.ParamSTB);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var list = conexaoBD.Query<GraficoComunicacaoVisto>("pr_Dashboard_ComunicacaoVisto", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    retorno = list;
                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }

        public List<GraficoComunicacaoVisto> CarregarGraficoComunicacaoSource(FiltroPadrao filtro)
        {
            var retorno = new List<GraficoComunicacaoVisto>();

            try
            {

                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoComunicacao(filtro);
                parametros.Add("@ParamSTB", filtro.ParamSTB);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var list = conexaoBD.Query<GraficoComunicacaoVisto>("pr_Dashboard_ComunicacaoSource", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    retorno = list;
                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }

        public List<GraficoComunicacaoDiagnostico> CarregarGraficoComunicacaoDiagnostico(FiltroPadrao filtro)
        {
            var retorno = new List<GraficoComunicacaoDiagnostico>();

            try
            {

                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroPadraoComunicacao(filtro);
                parametros.Add("@ParamSTB", filtro.ParamSTB);


                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var list = conexaoBD.Query<GraficoComunicacaoDiagnostico>("pr_Dashboard_ComunicacaoDiagnostico", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    retorno = list;
                }

            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;
        }

        public List<ComunicacaoQuadroResumo> CarregarComunicacaoQuadroResumo(FiltroPadrao filtro)
        {
            var retorno = new List<ComunicacaoQuadroResumo>();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroComunicacaoQuadroResumo(filtro);


                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    var list = conexaoBD.Query<ComunicacaoQuadroResumo>("pr_Dashboard_ComunicacaoQuadroResumo", parametros, null, false, 300, System.Data.CommandType.StoredProcedure).ToList();

                    retorno = list;
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
