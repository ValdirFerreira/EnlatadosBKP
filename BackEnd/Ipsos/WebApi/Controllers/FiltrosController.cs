using DataAccess.Filtros;
using DataAccess.LogUsuario;
using DataAccess.Traducao;
using DataAccess.Usuario;
using Entities.Filtros;
using Entities.LogUsuario;
using Entities.TraducaoIdioma;
using Entities.Usuario;
using Helpers.Logtxt;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Cors;
using WebApi.Models;




namespace WebApi.Controllers
{
    [EnableCors("*", "*", "*")]

    [RoutePrefix("api/filtros")]

    [Authorize]
    public class FiltrosController : ApiController
    {

        private FiltrosDataAccess _context = new FiltrosDataAccess(string.Empty);

        [HttpPost]
        [Route("FiltroTarget")]
        public HttpResponseMessage FiltroTarget(ParamGeralFiltro filtro)
        {
            var response = new Response();
            try
            {
                var dados = _context.FiltroTarget(filtro);

                return Request.CreateResponse(HttpStatusCode.OK, dados);
            }
            catch (Exception ex)
            {
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                response.Error = $"Bad request - ({ex.Message})";

                return Request.CreateResponse(HttpStatusCode.InternalServerError, response);
            }
        }

        [HttpPost]
        [Route("FiltroMarcasAdHoc")]
        public HttpResponseMessage FiltroMarcasAdHoc(ParamGeralFiltro filtro)
        {
            var response = new Response();
            try
            {
                var dados = _context.FiltroMarcasAdHoc(filtro);

                return Request.CreateResponse(HttpStatusCode.OK, dados);
            }
            catch (Exception ex)
            {
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                response.Error = $"Bad request - ({ex.Message})";

                return Request.CreateResponse(HttpStatusCode.InternalServerError, response);
            }
        }


        [HttpPost]
        [Route("FiltroMarcas")]
        public HttpResponseMessage FiltroMarcas(ParamGeralFiltro filtro)
        {
            var response = new Response();
            try
            {
                var dados = _context.FiltroMarcas(filtro);

                return Request.CreateResponse(HttpStatusCode.OK, dados);
            }
            catch (Exception ex)
            {
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                response.Error = $"Bad request - ({ex.Message})";

                return Request.CreateResponse(HttpStatusCode.InternalServerError, response);
            }
        }

        [HttpPost]
        [Route("FiltroOnda")]
        public HttpResponseMessage FiltroOnda(ParamGeralFiltro filtro)
        {
            var response = new Response();
            try
            {
                var dados = _context.FiltroOnda(filtro);

                return Request.CreateResponse(HttpStatusCode.OK, dados);
            }
            catch (Exception ex)
            {
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                response.Error = $"Bad request - ({ex.Message})";

                return Request.CreateResponse(HttpStatusCode.InternalServerError, response);
            }
        }

        [HttpPost]
        [Route("FiltroDemografico")]
        public HttpResponseMessage FiltroDemografico(ParamGeralFiltro filtro)
        {
            var response = new Response();
            try
            {
                var dados = _context.FiltroDemografico(filtro);

                return Request.CreateResponse(HttpStatusCode.OK, dados);
            }
            catch (Exception ex)
            {
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                response.Error = $"Bad request - ({ex.Message})";

                return Request.CreateResponse(HttpStatusCode.InternalServerError, response);
            }
        }

        [HttpPost]
        [Route("FiltroRegiao")]
        public HttpResponseMessage FiltroRegiao(ParamGeralFiltro filtro)
        {
            var response = new Response();
            try
            {
                var dados = _context.FiltroRegiao(filtro);

                return Request.CreateResponse(HttpStatusCode.OK, dados);
            }
            catch (Exception ex)
            {
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                response.Error = $"Bad request - ({ex.Message})";

                return Request.CreateResponse(HttpStatusCode.InternalServerError, response);
            }
        }


        [HttpPost]
        [Route("FiltroAtributos")]
        public HttpResponseMessage FiltroAtributos(ParamGeralFiltro filtro)
        {
            var response = new Response();
            try
            {
                var dados = _context.FiltroAtributos(filtro);

                return Request.CreateResponse(HttpStatusCode.OK, dados);
            }
            catch (Exception ex)
            {
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                response.Error = $"Bad request - ({ex.Message})";

                return Request.CreateResponse(HttpStatusCode.InternalServerError, response);
            }
        }


        [HttpPost]
        [Route("FiltroBVC")]
        public HttpResponseMessage FiltroBVC(ParamGeralFiltro filtro)
        {
            var response = new Response();
            try
            {
                var dados = _context.FiltroBVC(filtro);

                return Request.CreateResponse(HttpStatusCode.OK, dados);
            }
            catch (Exception ex)
            {
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                response.Error = $"Bad request - ({ex.Message})";

                return Request.CreateResponse(HttpStatusCode.InternalServerError, response);
            }
        }

        [HttpPost]
        [Route("FiltroSTB")]
        public HttpResponseMessage FiltroSTB(ParamGeralFiltro filtro)
        {
            var response = new Response();
            try
            {
                var dados = _context.FiltroSTB(filtro);

                return Request.CreateResponse(HttpStatusCode.OK, dados);
            }
            catch (Exception ex)
            {
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                response.Error = $"Bad request - ({ex.Message})";

                return Request.CreateResponse(HttpStatusCode.InternalServerError, response);
            }
        }

        [HttpPost]
        [Route("FiltroAdHoc")]
        public HttpResponseMessage FiltroAdHoc(ParamGeralFiltro filtro)
        {
            var response = new Response();
            try
            {
                var dados = _context.FiltroAdHoc();

                return Request.CreateResponse(HttpStatusCode.OK, dados);
            }
            catch (Exception ex)
            {
                response.StatusCode = (int)HttpStatusCode.InternalServerError;
                response.Error = $"Bad request - ({ex.Message})";

                return Request.CreateResponse(HttpStatusCode.InternalServerError, response);
            }
        }


        


    }
}
