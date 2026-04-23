using Business.Exportacoes;
using DataAccess.DashBoardEight;
using DataAccess.DashBoardFive;
using DataAccess.DashBoardFour;
using DataAccess.DashBoardNine;
using DataAccess.DashBoardSeven;
using DataAccess.DashBoardSix;
using DataAccess.DashboardThree;
using DataAccess.DashBoardTwo;
using DataAccess.Filtros;
using DataAccess.LogUsuario;
using DataAccess.Traducao;
using DataAccess.Usuario;
using Entities.Filtros;
using Entities.LogUsuario;
using Entities.Parametros;
using Entities.TraducaoIdioma;
using Entities.Usuario;
using Helpers.Logtxt;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Cors;
using WebApi.Models;




namespace WebApi.Controllers
{
    [EnableCors("*", "*", "*")]

    [RoutePrefix("api/excel")]
    [AllowAnonymous]
    public class ExcelController : ApiController
    {



        private TraducaoDataAccess _contextTraducao = new TraducaoDataAccess(Usuario.Email);
        private DashBoardTwoDataAccess _context = new DashBoardTwoDataAccess(Usuario.Email);
        private DashboardThreeDataAccess _contextDashboardThree = new DashboardThreeDataAccess(Usuario.Email);
        private DashBoardFourDataAccess _contextDashboardFour = new DashBoardFourDataAccess(Usuario.Email);
        private DashBoardFiveDataAccess _contextDashboardFive = new DashBoardFiveDataAccess(Usuario.Email);
        private DashBoardSixDataAccess _contextDashboardSix = new DashBoardSixDataAccess(Usuario.Email);
        private DashBoardSevenDataAccess _contextDashboardSeven = new DashBoardSevenDataAccess(Usuario.Email);
        private ExportaGraficoColuna _exportaExcelGraficoColuna = new ExportaGraficoColuna();
        private DashBoardEightDataAccess _contextDashBoardEight = new DashBoardEightDataAccess(Usuario.Email);
        private DashBoardNineDataAccess _contextDashboardNine = new DashBoardNineDataAccess(Usuario.Email);

        #region DownloadDashboardTwo

        [HttpPost]
        [Route("DownloadDashboardTwoComparativoMarcas")]

        public async Task<IHttpActionResult> DownloadDashboardTwoComparativoMarcas(FiltroPadraoExcel filtro)
        {
            try
            {
                MemoryStream result;
                var graficoColunasFullLoad = _context.CarregarGraficoComparativoMarcasExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardTwoComparativoMarcas(graficoColunasFullLoad, filtro.TituloGrafico, listTraducao));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }


        [HttpPost]
        [Route("DownloadDashboardTwoEvolutivoMarcas")]

        public async Task<IHttpActionResult> DownloadDashboardTwoEvolutivoMarcas(FiltroPadraoExcel filtro)
        {
            try
            {
                MemoryStream result;
                var graficoColunasFullLoad = _context.CarregarGraficoEvolutivoMarcasExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardTwoEvolutivoMarcas(graficoColunasFullLoad, filtro.TituloGrafico, listTraducao, filtro));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }


        [HttpPost]
        [Route("DownloadDashboardTwoComparativoMarcasDenominators")]

        public async Task<IHttpActionResult> DownloadDashboardTwoComparativoMarcasDenominators(FiltroPadraoExcel filtro)
        {
            try
            {
                MemoryStream result;
                var graficoColunasFullLoad = _context.CarregarGraficoComparativoMarcasExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardTwoComparativoMarcasDenominators(graficoColunasFullLoad, filtro.TituloGrafico, listTraducao));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }


        [HttpPost]
        [Route("DownloadDashboardTwoEvolutivoMarcasDenominators")]

        public async Task<IHttpActionResult> DownloadDashboardTwoEvolutivoMarcasDenominators(FiltroPadraoExcel filtro)
        {
            try
            {
                MemoryStream result;
                var graficoColunasFullLoad = _context.CarregarGraficoEvolutivoMarcasExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardTwoEvolutivoMarcasDenominators(graficoColunasFullLoad, filtro.TituloGrafico, listTraducao, filtro));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }

        #endregion


        #region DownloadDashboardThree

        [HttpPost]
        [Route("DownloadDashboardThreeComparativoMarcas")]

        public async Task<IHttpActionResult> DownloadDashboardThreeComparativoMarcas(FiltroPadraoExcel filtro)
        {
            try
            {
                MemoryStream result;
                var graficoColunasFullLoad = _contextDashboardThree.CarregarGraficoComparativoMarcasExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardThreeComparativoMarcas(graficoColunasFullLoad, filtro.TituloGrafico, listTraducao));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }


        [HttpPost]
        [Route("DownloadDashboardThreeEvolutivoMarcas")]

        public async Task<IHttpActionResult> DownloadDashboardThreeEvolutivoMarcas(FiltroPadraoExcel filtro)
        {
            try
            {
                MemoryStream result;
                var graficoColunasFullLoad = _contextDashboardThree.CarregarGraficoEvolutivoMarcasExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardThreeEvolutivoMarcas(graficoColunasFullLoad, filtro.TituloGrafico, listTraducao, filtro));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }

        #endregion


        #region DownloadDashboardFour

        [HttpPost]
        [Route("DownloadDashboardFourComparativoMarcas")]
        public async Task<IHttpActionResult> DownloadDashboardFourComparativoMarcas(FiltroPadrao filtro)
        {
            try
            {
                MemoryStream result;
                var graficoAranha = _contextDashboardFour.ComparativoMarcasExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardFourComparativoMarcas(graficoAranha, filtro.TituloGrafico, listTraducao));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }


        [HttpPost]
        [Route("DownloadDashboardFourEvolutivoPeriodos")]
        public async Task<IHttpActionResult> DownloadDashboardFourEvolutivoPeriodos(FiltroPadraoExcel filtro)
        {
            try
            {
                MemoryStream result;
                var graficoAranha = _contextDashboardFour.CarregarGraficoImagemEvolutivaExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardFourGraficoLinhasImagem(graficoAranha, filtro.TituloGrafico, listTraducao, filtro));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }

        [HttpPost]
        [Route("DownloadDashboardFourComparativoImagemPura")]

        public async Task<IHttpActionResult> DownloadDashboardFourComparativoImagemPura(FiltroPadrao filtro)
        {
            try
            {
                MemoryStream result;
                var graficoAranha = _contextDashboardFour.ImagemEvolutivaLinhasExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardFourImagemPura(graficoAranha, filtro.TituloGrafico, listTraducao));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }


        #endregion


        #region DownloadDashboardFive

        [HttpPost]
        [Route("DownloadDashboardFiveComparativoMarcas")]

        public async Task<IHttpActionResult> DownloadDashboardFiveComparativoMarcas(FiltroPadraoExcel filtro)
        {
            try
            {
                MemoryStream result;
                var graficoColunasFullLoad = _contextDashboardFive.CarregarGraficoComparativoMarcasExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardFiveComparativoMarcas(graficoColunasFullLoad, filtro.TituloGrafico, listTraducao));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }

        [HttpPost]
        [Route("DownloadDashboardFiveEvolutivoMarcas")]

        public async Task<IHttpActionResult> DownloadDashboardFiveEvolutivoMarcas(FiltroPadraoExcel filtro)
        {
            try
            {
                MemoryStream result;
                var graficoColunasFullLoad = _contextDashboardFive.CarregarGraficoEvolutivoMarcasExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardFiveEvolutivoMarcas(graficoColunasFullLoad, filtro.TituloGrafico, listTraducao, filtro));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }


        [HttpPost]
        [Route("DownloadDashboardFiveComparativoMarcasDuplo")]

        public async Task<IHttpActionResult> DownloadDashboardFiveComparativoMarcasDuplo(FiltroPadraoExcelDuplo filtro)
        {
            try
            {
                MemoryStream result;
                var graficoColunasFullLoad = _contextDashboardFive.CarregarGraficoEvolutivoMarcasDuploExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardFiveComparativoDuploMarcasNew(graficoColunasFullLoad, filtro.TituloGrafico, listTraducao, filtro));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }

        #endregion


        #region DownloadDashboardSix



        [HttpPost]
        [Route("DownloadDashboardSixGraficoBVCTop10Marcas")]

        public async Task<IHttpActionResult> DownloadDashboardSixGraficoBVCTop10Marcas(FiltroPadrao filtro)
        {
            try
            {
                MemoryStream result;
                var grafico = _contextDashboardSix.CarregarGraficoBVCTop10Marcas(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardSixGraficoBVCTop10Marcas(grafico, filtro.TituloGrafico, listTraducao));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }


        [HttpPost]
        [Route("DownloadDashboardSixGraficoBVCEvolutivo")]

        public async Task<IHttpActionResult> DownloadDashboardSixGraficoBVCEvolutivo(FiltroPadraoExcel filtro)
        {
            try
            {
                MemoryStream result;
                var grafico = _contextDashboardSix.CarregarGraficoBVCEvolutivoExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardSixGraficoBVCEvolutivo(grafico, filtro.TituloGrafico, listTraducao, filtro));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }

        [HttpPost]
        [Route("DownloadDashboardSixComparativoMarcas")]

        public async Task<IHttpActionResult> DownloadDashboardSixComparativoMarcas(FiltroPadraoExcel filtro)
        {
            try
            {
                MemoryStream result;
                var graficoColunasFullLoad = _contextDashboardSix.CarregarGraficoComparativoMarcasExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardFiveComparativoMarcas(graficoColunasFullLoad, filtro.TituloGrafico, listTraducao));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }

        [HttpPost]
        [Route("DownloadDashboardSixEvolutivoMarcas")]

        public async Task<IHttpActionResult> DownloadDashboardSixEvolutivoMarcas(FiltroPadraoExcel filtro)
        {
            try
            {
                MemoryStream result;
                var graficoColunasFullLoad = _contextDashboardSix.CarregarGraficoEvolutivoMarcasExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardFiveEvolutivoMarcas(graficoColunasFullLoad, filtro.TituloGrafico, listTraducao, filtro));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }


        [HttpPost]
        [Route("DownloadDashboardSixComparativoMarcasDuplo")]

        public async Task<IHttpActionResult> DownloadDashboardSixComparativoMarcasDuplo(FiltroPadraoExcelDuplo filtro)
        {
            try
            {
                MemoryStream result;
                var graficoColunasFullLoad = _contextDashboardSix.CarregarGraficoEvolutivoMarcasDuploExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardFiveComparativoDuploMarcas(graficoColunasFullLoad, filtro.TituloGrafico, listTraducao, filtro));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }

        #endregion


        #region DownloadDashboardSeven

        [HttpPost]
        [Route("DownloadDashboardGraficoComunicacaoRecall")]

        public async Task<IHttpActionResult> DownloadDashboardGraficoComunicacaoRecall(FiltroPadrao filtro)
        {
            try
            {
                MemoryStream result;
                var grafico = _contextDashboardSeven.CarregarGraficoComunicacaoRecall(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardGraficoComunicacaoRecall(grafico, filtro.TituloGrafico, listTraducao));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }


        [HttpPost]
        [Route("DownloadDashboardGraficoComunicacaoVisto")]

        public async Task<IHttpActionResult> DownloadDashboardGraficoComunicacaoVisto(FiltroPadrao filtro)
        {
            try
            {
                MemoryStream result;
                var grafico = _contextDashboardSeven.CarregarGraficoComunicacaoVisto(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardGraficoComunicacaoVisto(grafico, filtro.TituloGrafico, listTraducao));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }


        [HttpPost]
        [Route("DownloadDashboardGraficoComunicacaoSource")]

        public async Task<IHttpActionResult> DownloadDashboardGraficoComunicacaoSource(FiltroPadrao filtro)
        {
            try
            {
                MemoryStream result;
                var grafico = _contextDashboardSeven.CarregarGraficoComunicacaoSource(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardGraficoComunicacaoSource(grafico, filtro.TituloGrafico, listTraducao));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }


        [Route("DownloadDashboardGraficoComunicacaoDiagnostico")]
        public async Task<IHttpActionResult> DownloadDashboardGraficoComunicacaoDiagnostico(FiltroPadrao filtro)
        {
            try
            {
                MemoryStream result;
                var grafico = _contextDashboardSeven.CarregarGraficoComunicacaoDiagnostico(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardGraficoComunicacaoDiagnostico(grafico, filtro.TituloGrafico, listTraducao));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }

        #endregion


        #region DownloadDashboardEight

        [HttpPost]
        [Route("DownloadDashboardEightComparativoMarcas")]

        public async Task<IHttpActionResult> DownloadDashboardEightComparativoMarcas(FiltroPadraoExcel filtro)
        {
            try
            {
                MemoryStream result;
                var graficoColunasFullLoad = _contextDashBoardEight.CarregarGraficoComparativoMarcasConsiderationExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardEightComparativoMarcas(graficoColunasFullLoad, filtro.TituloGrafico, listTraducao));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }


        [HttpPost]
        [Route("DownloadDashboardEightEvolutivoMarcas")]

        public async Task<IHttpActionResult> DownloadDashboardEightEvolutivoMarcas(FiltroPadraoExcel filtro)
        {
            try
            {
                MemoryStream result;
                var graficoColunasFullLoad = _contextDashBoardEight.CarregarGraficoEvolutivoMarcasConsiderationExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardEightEvolutivoMarcas(graficoColunasFullLoad, filtro.TituloGrafico, listTraducao, filtro));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }

        #endregion



        #region DownloadDashboardNine
       
        [HttpPost]
        [Route("DownloadDashboardNineEvolutivoPeriodos")]
        public async Task<IHttpActionResult> DownloadDashboardNineEvolutivoPeriodos(FiltroPadraoExcel filtro)
        {
            try
            {
                MemoryStream result;
                var graficoAranha = _contextDashboardNine.CarregarGraficoImagemEvolutivaNineExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardNineGraficoLinhasImagem(graficoAranha, filtro.TituloGrafico, listTraducao, filtro));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }

        [HttpPost]
        [Route("DownloadDashboardNineComparativoImagemPura")]

        public async Task<IHttpActionResult> DownloadDashboardNineComparativoImagemPura(FiltroPadrao filtro)
        {
            try
            {
                MemoryStream result;
                var graficoAranha = _contextDashboardNine.ImagemEvolutivaLinhasNineExcel(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardNineImagemPura(graficoAranha, filtro.TituloGrafico, listTraducao));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }


        [HttpPost]
        [Route("DownloadDashboardNineComparativoImagemPura2")]

        public async Task<IHttpActionResult> DownloadDashboardNineComparativoImagemPura2(FiltroPadrao filtro)
        {
            try
            {
                MemoryStream result;
                var graficoAranha = _contextDashboardNine.ImagemEvolutivaLinhasNineExcel2(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();
                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardNineImagemPura2(graficoAranha, filtro.TituloGrafico, listTraducao));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }


        [HttpPost]
        [Route("DownloadDashboardNineTabelaAdHoc")]

        public async Task<IHttpActionResult> DownloadDashboardNineTabelaAdHoc(FiltroPadrao filtro)
        {
            try
            {
                MemoryStream result;
                var graficoAranha = _contextDashboardNine.TabelaAdHocAtributo(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();

                var marca = filtro.Marca.FirstOrDefault();;

                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardNineTabelaAdHocAtributo(graficoAranha, filtro.TituloGrafico, listTraducao, marca));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }



        [HttpPost]
        [Route("DownloadDashboardNineTabelaAdHocBloco6")]

        public async Task<IHttpActionResult> DownloadDashboardNineTabelaAdHocBloco6(FiltroPadrao filtro)
        {
            try
            {
                MemoryStream result;
                var graficoAranha = _contextDashboardNine.TabelaAdHocAtributoBloco6(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();

                var marca = filtro.Marca.FirstOrDefault(); ;

                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardNineTabelaAdHocAtributo(graficoAranha, filtro.TituloGrafico, listTraducao, null));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }



        [HttpPost]
        [Route("DownloadDashboardNineTabelaAdHocBloco10")]

        public async Task<IHttpActionResult> DownloadDashboardNineTabelaAdHocBloco10(FiltroPadrao filtro)
        {
            try
            {
                MemoryStream result;
                var graficoAranha = _contextDashboardNine.TabelaAdHocAtributoBloco10(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();

                var marca = filtro.Marca.FirstOrDefault(); ;

                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardNineTabelaAdHocAtributo(graficoAranha, filtro.TituloGrafico, listTraducao, marca));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }


        [HttpPost]
        [Route("DownloadDashboardNineTabelaAdHocBloco2")]

        public async Task<IHttpActionResult> DownloadDashboardNineTabelaAdHocBloco2(FiltroPadrao filtro)
        {
            try
            {
                MemoryStream result;
                var graficoAranha = _contextDashboardNine.TabelaAdHocAtributoBloco2(filtro);

                var listTraducao = _contextTraducao.ObtemTraducoes();

                var marca = filtro.Marca !=null ? filtro.Marca.FirstOrDefault(): null; 

                result = new MemoryStream(_exportaExcelGraficoColuna.DownloadDashboardNineTabelaAdHocAtributo(graficoAranha, filtro.TituloGrafico, listTraducao, marca));
                return new ArquivoResult(result, Request, "Arquivo" + " " + DateTime.Now.ToString("yyyy-MM-dd HH':'mm':'ss':'fff").Replace("-", "").Replace(":", "").Replace(" ", "") + ".xlsx");
            }
            catch (Exception ex)
            {
                return ResponseMessage(Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message));
            }
        }



        #endregion


    }
}

