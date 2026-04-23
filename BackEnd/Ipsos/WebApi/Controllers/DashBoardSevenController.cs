using DataAccess.DashBoardFive;
using DataAccess.DashBoardSeven;
using DataAccess.DashBoardSix;
using DataAccess.DashBoardTwo;
using DataAccess.Traducao;
using Entities.Parametros;
using Helpers.Logtxt;
using System.Data.SqlClient;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Cors;
using WebApi.Models;




namespace WebApi.Controllers
{
    [EnableCors("*", "*", "*")]
    [RoutePrefix("api/DashBoardSeven")]


    [Authorize]
    public class DashBoardSevenController : ApiController
    {

        private DashBoardSevenDataAccess _context = new DashBoardSevenDataAccess(Usuario.Email);



        [HttpPost]
        [Route("CarregarGraficoComunicacaoRecall")]
        public HttpResponseMessage CarregarGraficoComunicacaoRecall(FiltroPadrao filtro)
        {
            var response = new Response();
            try
            {
                var list = _context.CarregarGraficoComunicacaoRecall(filtro);

                return Request.CreateResponse(HttpStatusCode.OK, list);

            }
            catch (SqlException ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "Sistema" + ex.Message);
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                response.Error = $"Bad request - ({ex.Message})";
                return Request.CreateResponse(HttpStatusCode.InternalServerError, response);
            }
        }


        [HttpPost]
        [Route("CarregarGraficoComunicacaoVisto")]
        public HttpResponseMessage CarregarGraficoComunicacaoVisto(FiltroPadrao filtro)
        {
            var response = new Response();
            try
            {
               var list =  _context.CarregarGraficoComunicacaoVisto(filtro);

                return Request.CreateResponse(HttpStatusCode.OK, list);

            }
            catch (SqlException ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "Sistema" + ex.Message);
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                response.Error = $"Bad request - ({ex.Message})";
                return Request.CreateResponse(HttpStatusCode.InternalServerError, response);
            }
        }


        [HttpPost]
        [Route("CarregarGraficoComunicacaoSource")]
        public HttpResponseMessage CarregarGraficoComunicacaoSource(FiltroPadrao filtro)
        {
            var response = new Response();
            try
            {
                var list = _context.CarregarGraficoComunicacaoSource(filtro);

                return Request.CreateResponse(HttpStatusCode.OK, list);

            }
            catch (SqlException ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "Sistema" + ex.Message);
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                response.Error = $"Bad request - ({ex.Message})";
                return Request.CreateResponse(HttpStatusCode.InternalServerError, response);
            }
        }

        [HttpPost]
        [Route("CarregarGraficoComunicacaoDiagnostico")]
        public HttpResponseMessage CarregarGraficoComunicacaoDiagnostico(FiltroPadrao filtro)
        {
            var response = new Response();
            try
            {
                var list = _context.CarregarGraficoComunicacaoDiagnostico(filtro);

                return Request.CreateResponse(HttpStatusCode.OK, list);

            }
            catch (SqlException ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "Sistema" + ex.Message);
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                response.Error = $"Bad request - ({ex.Message})";
                return Request.CreateResponse(HttpStatusCode.InternalServerError, response);
            }
        }


        [HttpPost]
        [Route("CarregarComunicacaoQuadroResumo")]
        public HttpResponseMessage CarregarComunicacaoQuadroResumo(FiltroPadrao filtro)
        {
            var response = new Response();
            try
            {
                var list = _context.CarregarComunicacaoQuadroResumo(filtro);

                return Request.CreateResponse(HttpStatusCode.OK, list);

            }
            catch (SqlException ex)
            {
                LogText.Instance.Error(this.GetType().Name, System.Reflection.MethodBase.GetCurrentMethod().Name, "Sistema" + ex.Message);
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                response.Error = $"Bad request - ({ex.Message})";
                return Request.CreateResponse(HttpStatusCode.InternalServerError, response);
            }
        }

    }
}
