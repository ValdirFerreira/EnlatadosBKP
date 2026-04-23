using DataAccess.DashboardThree;
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
    [RoutePrefix("api/DashBoardThree")]


    [Authorize]
    public class DashBoardThreeController : ApiController
    {

        private DashboardThreeDataAccess _context = new DashboardThreeDataAccess(Usuario.Email);

       


        [HttpPost]
        [Route("CarregarGraficoComparativoExperiencia")]
        public HttpResponseMessage CarregarGraficoComparativoExperiencia(FiltroPadrao filtro)
        {
            var response = new Response();
            try
            {
               var list =  _context.CarregarGraficoComparativoExperiencia(filtro);

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
