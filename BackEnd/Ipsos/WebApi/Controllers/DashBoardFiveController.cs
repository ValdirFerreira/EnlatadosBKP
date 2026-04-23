using DataAccess.DashBoardFive;
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
    [RoutePrefix("api/DashBoardFive")]

    [Authorize]
    public class DashBoardFiveController : ApiController
    {

        private DashBoardFiveDataAccess _context = new DashBoardFiveDataAccess(Usuario.Email);




        [HttpPost]
        [Route("CarregarGraficoComparativoMarcas")]
        public HttpResponseMessage CarregarGraficoComparativoMarcas(FiltroPadrao filtro)
        {
            var response = new Response();
            try
            {
                var list = _context.CarregarGraficoComparativoMarcas(filtro);

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
        [Route("CarregarGraficoComparativoMarcas2")]
        public HttpResponseMessage CarregarGraficoComparativoMarcas2(FiltroPadrao filtro)
        {
            var response = new Response();
            try
            {
                var list = _context.CarregarGraficoComparativoMarcas2(filtro);

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
