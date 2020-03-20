using com.microsoft.dx.officewopi.Models.Wopi;
using com.microsoft.dx.officewopi.Utils;
using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace com.microsoft.dx.officewopi.Controllers
{
    //[WopiTokenValidationFilter]
    public class WopiController : ApiController
    {

        //[WopiTokenValidationFilter]
        [HttpGet]
        [Route("wopi/files/{id}/contents")]
        public async Task<HttpResponseMessage> Contents(Guid id)
        {
            //Handles GetFile
            return await WopiRequestHandler.ProcessWopiRequest(HttpContext.Current, id, WopiRequestType.GetFile);
        }

        //[WopiTokenValidationFilter]
        [HttpPost]
        [Route("wopi/files/{id}/contents")]
        public async Task<HttpResponseMessage> PostContents(Guid id)
        {
            //Handles PutFile
            return await WopiRequestHandler.ProcessWopiRequest(HttpContext.Current, id, WopiRequestType.PutFile);
        }

        //[WopiTokenValidationFilter]
        [HttpGet]
        [Route("wopi/files/{id}")]
        public async Task<HttpResponseMessage> Get(Guid id)
        {
            //Handles CheckFileInfo
            var result = await WopiRequestHandler.ProcessWopiRequest(HttpContext.Current, id, WopiRequestType.CheckFileInfo);
            return result;
        }

        //[WopiTokenValidationFilter]
        [HttpPost]
        [Route("wopi/files/{id}")]
        public async Task<HttpResponseMessage> Post(Guid id)
        {
            var requestType = WopiRequestType.None;

            //Handles Lock, GetLock, RefreshLock, Unlock, UnlockAndRelock, PutRelativeFile, RenameFile, PutUserInfo
            string wopiOverride = HttpContext.Current.Request.Headers[WopiRequestHeaders.OVERRIDE];

            switch (wopiOverride)
            {
                case "LOCK":
                    requestType = HttpContext.Current.Request.Headers[WopiRequestHeaders.OLD_LOCK] == null ? 
                                    WopiRequestType.Lock : WopiRequestType.UnlockAndRelock;
                    break;
                case "GET_LOCK":
                    requestType = WopiRequestType.GetLock;
                    break;
                case "REFRESH_LOCK":
                    requestType = WopiRequestType.RefreshLock;
                    break;
                case "UNLOCK":
                    requestType = WopiRequestType.Unlock;
                    break;
                case "PUT_RELATIVE":
                    requestType = WopiRequestType.PutRelativeFile;
                    break;
                case "RENAME_FILE":
                    requestType = WopiRequestType.RenameFile;
                    break;
                case "PUT_USER_INFO":
                    requestType = WopiRequestType.PutUserInfo;
                    break;
            }

            return await WopiRequestHandler.ProcessWopiRequest(HttpContext.Current, id, requestType);
        }

    }
}
