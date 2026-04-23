using DataAccess.DashBoardFour;
using DataAccess.Traducao;
using Entities.Filtros;
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
    [RoutePrefix("api/DashBoardFour")]


    [Authorize]
    public class DashBoardFourController : ApiController
    {

        private DashBoardFourDataAccess _context = new DashBoardFourDataAccess(Usuario.Email);

        [HttpPost]
        [Route("ComparativoMarcas")]
        public HttpResponseMessage ComparativoMarcas(FiltroPadrao filtro)
        {
            var response = new Response();
            try
            {
               var list =  _context.ComparativoMarcas(filtro);

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
        [Route("ImagemEvolutiva")]
        public HttpResponseMessage ImagemEvolutiva(FiltroPadrao filtro)
        {
            var response = new Response();
            try
            {
                var list = _context.ImagemEvolutiva(filtro);

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
        [Route("ImagemEvolutivaLinhas")]
        public HttpResponseMessage ImagemEvolutivaLinhas(FiltroPadrao filtro)
        {
            var response = new Response();
            try
            {
                var list = _context.ImagemEvolutivaLinhas(filtro);

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
