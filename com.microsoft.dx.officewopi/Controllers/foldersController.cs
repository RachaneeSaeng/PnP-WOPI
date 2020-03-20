using com.microsoft.dx.officewopi.Utils;
using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace com.microsoft.dx.officewopi.Controllers
{
    //[WopiTokenValidationFilter]
    public class foldersController : ApiController
    {
        //[WopiTokenValidationFilter]
        [HttpGet]
        [Route("wopi/folders/{id}")]
        public async Task<HttpResponseMessage> Get(Guid id)
        {
            return await WopiRequestHandler.ProcessWopiRequest(HttpContext.Current);
        }

        //[WopiTokenValidationFilter]
        [HttpGet]
        [Route("wopi/folders/{id}/contents")]
        public async Task<HttpResponseMessage> Contents(Guid id)
        {
            return await WopiRequestHandler.ProcessWopiRequest(HttpContext.Current);
        }

        //[WopiTokenValidationFilter]
        [HttpPost]
        [Route("wopi/folders/{id}")]
        public async Task<HttpResponseMessage> Post(Guid id)
        {
            return await WopiRequestHandler.ProcessWopiRequest(HttpContext.Current);
        }

        //[WopiTokenValidationFilter]
        [HttpPost]
        [Route("wopi/folders/{id}/contents")]
        public async Task<HttpResponseMessage> PostContents(Guid id)
        {
            return await WopiRequestHandler.ProcessWopiRequest(HttpContext.Current);
        }
    }
}
