using com.microsoft.dx.officewopi.Models;
using com.microsoft.dx.officewopi.Models.Wopi;
using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;

namespace com.microsoft.dx.officewopi.Utils
{
    public static class WopiRequestHandler
    {
        /// <summary>
        /// Processes a WOPI request using the HttpContext of the APIController
        /// </summary>
        public async static Task<HttpResponseMessage> ProcessWopiRequest(HttpContext context, Guid id, WopiRequestType requestType)
        {
            try
            {
                // Validate WOPI Proof (ie - ensure request came from Office Online)
                if (!(await WopiRequestUtil.ValidateWopiProof(context)))
                    return ReturnStatus(HttpStatusCode.InternalServerError, "Server Error");

                // Lookup the file in the database
                var file = DocumentDBRepository<DetailedFileModel>.GetItem("Files", i => i.id == id);

                // Check for null file
                if (file == null)
                    return ReturnStatus(HttpStatusCode.NotFound, "File Unknown/User Unauthorized");

                // Get discovery information
                await file.PopulateActions();

                // Augments the file with additional properties CloseUrl, HostViewUrl, HostEditUrl
                file.CloseUrl = String.Format("https://{0}", context.Request.Url.Authority);

                var view = file.Actions.FirstOrDefault(i => i.name == "view");
                if (view != null)
                    file.HostViewUrl = $"https://wopi-test.go.myworkpapers.co.uk/Home/Detail/{file.id}?action=view";
                //file.HostViewUrl = WopiUtil.GetActionUrl(view, file, context.Request.Url.Authority);
                var edit = file.Actions.FirstOrDefault(i => i.name == "edit");
                if (edit != null)
                    file.HostEditUrl = $"https://wopi-test.go.myworkpapers.co.uk/Home/Detail/{file.id}?action=edit";
                //file.HostEditUrl = WopiUtil.GetActionUrl(edit, file, context.Request.Url.Authority);

                // Get the user from the token (token is already validated)
                file.UserId = "rachanee.saeng@gmail.com";
                file.UserFriendlyName = "Apple Saeng";

                // Call the appropriate handler for the WOPI request we received
                switch (requestType)
                {
                    case WopiRequestType.CheckFileInfo:
                        return context.CheckFileInfo(file);
                    case WopiRequestType.GetFile:
                        return await context.GetFile(file);
                    case WopiRequestType.Lock:
                        return await context.Lock(file);
                    case WopiRequestType.GetLock:
                        return await context.GetLock(file);
                    case WopiRequestType.RefreshLock:
                        return await context.RefreshLock(file);
                    case WopiRequestType.Unlock:
                        return await context.Unlock(file);
                    case WopiRequestType.UnlockAndRelock:
                        return await context.UnlockAndRelock(file);
                    case WopiRequestType.PutFile:
                        return await context.PutFile(file);
                    case WopiRequestType.PutRelativeFile:
                        return await context.PutRelativeFile(file);
                    case WopiRequestType.RenameFile:
                        return await context.RenameFile(file);
                    case WopiRequestType.PutUserInfo:
                        return await context.PutUserInfo(file);
                    default:
                        return ReturnStatus(HttpStatusCode.NotImplemented, "Unsupported");
                }

            }
            catch (Exception)
            {
                return ReturnStatus(HttpStatusCode.InternalServerError, "Server Error");
            }
        }

        /// <summary>
        /// Forms the HttpResponseMessage for the WOPI request
        /// </summary>
        private static HttpResponseMessage ReturnStatus(HttpStatusCode code, string description)
        {
            HttpResponseMessage response = new HttpResponseMessage(code);
            response.ReasonPhrase = description;
            return response;
        }
    }
}