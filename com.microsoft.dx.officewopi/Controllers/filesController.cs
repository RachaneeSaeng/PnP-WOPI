using com.microsoft.dx.officewopi.Utils;
using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace com.microsoft.dx.officewopi.Controllers
{
    //[WopiTokenValidationFilter]
    public class filesController : ApiController
    {
        //[WopiTokenValidationFilter]
        [HttpGet]
        [Route("wopi/files/{id}")]
        public async Task<HttpResponseMessage> Get(Guid id)
        {
            //Handles CheckFileInfo
            var result = await WopiRequestHandler.ProcessWopiRequest(HttpContext.Current);
            return result;
        }

        //[WopiTokenValidationFilter]
        [HttpGet]
        [Route("wopi/files/{id}/contents")]
        public async Task<HttpResponseMessage> Contents(Guid id)
        {
            //Handles GetFile
            return await WopiRequestHandler.ProcessWopiRequest(HttpContext.Current);
        }

        //[WopiTokenValidationFilter]
        [HttpPost]
        [Route("wopi/files/{id}")]
        public async Task<HttpResponseMessage> Post(Guid id)
        {
            //Handles Lock, GetLock, RefreshLock, Unlock, UnlockAndRelock, PutRelativeFile, RenameFile, PutUserInfo
            return await WopiRequestHandler.ProcessWopiRequest(HttpContext.Current);
        }

        //[WopiTokenValidationFilter]
        [HttpPost]
        [Route("wopi/files/{id}/contents")]
        public async Task<HttpResponseMessage> PostContents(Guid id)
        {
            //Handles PutFile
            return await WopiRequestHandler.ProcessWopiRequest(HttpContext.Current);
        }
    }
}
