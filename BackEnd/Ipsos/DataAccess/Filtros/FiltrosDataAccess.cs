using Business.TrataDados;
using Dapper;
using DataAccess.Config;
using Entities.Filtros;
using Helpers.Logtxt;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess.Filtros
{
    public class FiltrosDataAccess
    {
        private readonly string usuarioEmail = string.Empty;

        public FiltrosDataAccess(string usuario)
        {
            usuarioEmail = usuario;
        }

        public List<PadraoComboFiltro> FiltroTarget(ParamGeralFiltro filtro)
        {
            var retorno = new List<PadraoComboFiltro>();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltros(filtro);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    retorno = conexaoBD.Query<PadraoComboFiltro>("pr_FiltroTarget", parametros, null, true, 300, System.Data.CommandType.StoredProcedure).ToList();
                }
            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }

        public List<PadraoComboFiltro> FiltroMarcas(ParamGeralFiltro filtro)
        {
            var retorno = new List<PadraoComboFiltro>();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltrosMarca(filtro);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    retorno = conexaoBD.Query<PadraoComboFiltro>("pr_FiltroMarcas", parametros, null, true, 300, System.Data.CommandType.StoredProcedure).ToList();

                }
            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }

        public List<PadraoComboFiltro> FiltroMarcasAdHoc(ParamGeralFiltro filtro)
        {
            var retorno = new List<PadraoComboFiltro>();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltros(filtro);
                parametros.Add("@CodBlocoParam", filtro.CodBlocoParam);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    retorno = conexaoBD.Query<PadraoComboFiltro>("pr_FiltroMarcasAdHoc", parametros, null, true, 300, System.Data.CommandType.StoredProcedure).ToList();

                }
            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }

        public List<PadraoComboFiltro> FiltroOnda(ParamGeralFiltro filtro)
        {
            var retorno = new List<PadraoComboFiltro>();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltros(filtro);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    retorno = conexaoBD.Query<PadraoComboFiltro>("pr_FiltroOnda", parametros, null, true, 300, System.Data.CommandType.StoredProcedure).ToList();

                }
            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }

        public List<PadraoComboFiltro> FiltroDemografico(ParamGeralFiltro filtro)
        {
            var retorno = new List<PadraoComboFiltro>();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltros(filtro);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    retorno = conexaoBD.Query<PadraoComboFiltro>("pr_FiltroDemografico", parametros, null, true, 300, System.Data.CommandType.StoredProcedure).ToList();

                }
            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }

        public List<PadraoComboFiltro> FiltroRegiao(ParamGeralFiltro filtro)
        {
            var retorno = new List<PadraoComboFiltro>();

            try
            {
                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltros(filtro);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    retorno = conexaoBD.Query<PadraoComboFiltro>("pr_FiltroRegiao", parametros, null, true, 300, System.Data.CommandType.StoredProcedure).ToList();

                }
            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }


        public List<PadraoComboFiltro> FiltroAtributos(ParamGeralFiltro filtro)
        {
            var retorno = new List<PadraoComboFiltro>();

            try
            {
                //var parametros = TrataFiltros.MontaParametrosFiltros(filtro);

                var parametros = new DynamicParameters();
                parametros.Add("@ParamBIA", filtro.ParamBIA);
                parametros.Add("@ParamOnda", filtro.CodOndaParam);
                //parametros.Add("@ParamTarget", filtro.ParamTarget);



                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    retorno = conexaoBD.Query<PadraoComboFiltro>("pr_FiltroAtributos", parametros, null, true, 300, System.Data.CommandType.StoredProcedure).ToList();
                }
            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }



        public List<PadraoComboFiltro> FiltroBVC(ParamGeralFiltro filtro)
        {
            var retorno = new List<PadraoComboFiltro>();

            try
            {

                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltros(filtro);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    retorno = conexaoBD.Query<PadraoComboFiltro>("pr_FiltroBVC", parametros, null, true, 300, System.Data.CommandType.StoredProcedure).ToList();

                }
            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }


        public List<PadraoComboFiltro> FiltroSTB(ParamGeralFiltro filtro)
        {
            var retorno = new List<PadraoComboFiltro>();

            try
            {

                var TrataFiltros = new TrataFiltros();
                var parametros = TrataFiltros.MontaParametrosFiltroSTB(filtro);

                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    retorno = conexaoBD.Query<PadraoComboFiltro>("pr_FiltroSTB", parametros, null, true, 300, System.Data.CommandType.StoredProcedure).ToList();

                }
            }
            catch (Exception ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "[" + usuarioEmail + "]" + ex.Message);
            }

            return retorno;

        }


        public List<PadraoComboFiltro> FiltroAdHoc()
        {
            var retorno = new List<PadraoComboFiltro>();

            try
            {
                //var TrataFiltros = new TrataFiltros();
                //var parametros = TrataFiltros.MontaParametrosFiltros(filtro);


                using (SqlConnection conexaoBD = new SqlConnection(Conexao.strConexao))
                {
                    retorno = conexaoBD.Query<PadraoComboFiltro>("pr_FiltroAdHoc", null, null, true, 300, System.Data.CommandType.StoredProcedure).ToList();

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
