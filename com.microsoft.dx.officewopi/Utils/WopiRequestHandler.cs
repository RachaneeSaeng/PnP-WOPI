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
        public async static Task<HttpResponseMessage> ProcessWopiRequest(HttpContext context)
        {
            // Parse the request
            var request = ParseRequest(context.Request);
            HttpResponseMessage response = null;

            try
            {
                // Lookup the file in the database
                var itemId = new Guid(request.Id);
                var file = DocumentDBRepository<DetailedFileModel>.GetItem("Files", i => i.id == itemId);

                // Check for null file
                if (file == null)
                    response = ReturnStatus(HttpStatusCode.NotFound, "File Unknown/User Unauthorized");
                else
                {
                    // Validate WOPI Proof (ie - ensure request came from Office Online)
                    if (await WopiRequestUtil.ValidateWopiProof(context))
                    {
                        // Get discovery information
                        var fileExt = file.BaseFileName.Substring(file.BaseFileName.LastIndexOf('.') + 1).ToLower();
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
                        switch (request.RequestType)
                        {
                            case WopiRequestType.CheckFileInfo:
                                response = context.CheckFileInfo(file);
                                break;
                            case WopiRequestType.GetFile:
                                response = await context.GetFile(file);
                                break;
                            case WopiRequestType.Lock:
                                response = await context.Lock(file);
                                break;
                            case WopiRequestType.GetLock:
                                response = await context.GetLock(file);
                                break;
                            case WopiRequestType.RefreshLock:
                                response = await context.RefreshLock(file);
                                break;
                            case WopiRequestType.Unlock:
                                response = await context.Unlock(file);
                                break;
                            case WopiRequestType.UnlockAndRelock:
                                response = await context.UnlockAndRelock(file);
                                break;
                            case WopiRequestType.PutFile:
                                response = await context.PutFile(file);
                                break;
                            case WopiRequestType.PutRelativeFile:
                                response = await context.PutRelativeFile(file);
                                break;
                            case WopiRequestType.RenameFile:
                                response = await context.RenameFile(file);
                                break;
                            case WopiRequestType.PutUserInfo:
                                response = await context.PutUserInfo(file);
                                break;
                            default:
                                response = ReturnStatus(HttpStatusCode.NotImplemented, "Unsupported");
                                break;
                        }
                    }
                    else
                    {
                        // Proof validation failed...return 500
                        response = ReturnStatus(HttpStatusCode.InternalServerError, "Server Error");
                    }
                }
            }
            catch (Exception)
            {
                // An unknown exception occurred...return 500
                response = ReturnStatus(HttpStatusCode.InternalServerError, "Server Error");
            }

            return response;
        }


        /// <summary>
        /// Called at the beginning of a WOPI request to parse the request and determine the request type
        /// </summary>
        private static WopiRequest ParseRequest(HttpRequest request)
        {
            // Initilize wopi request data object with default values
            WopiRequest requestData = new WopiRequest()
            {
                RequestType = WopiRequestType.None,
                AccessToken = request.QueryString["access_token"],
                Id = ""
            };

            // Get request path, e.g. /<...>/wopi/files/<id>
            string requestPath = request.Url.AbsolutePath.ToLower();

            // Remove /<...>/wopi/
            string wopiPath = requestPath.Substring(requestPath.IndexOf(WopiRequestConsts.WOPI_BASE_PATH) + WopiRequestConsts.WOPI_BASE_PATH.Length);

            // Check the type of request being made
            if (wopiPath.StartsWith(WopiRequestConsts.WOPI_FILES_PATH))
            {
                // This is a file-related request

                // Remove /files/ from the beginning of wopiPath
                string rawId = wopiPath.Substring(WopiRequestConsts.WOPI_FILES_PATH.Length);

                if (rawId.EndsWith(WopiRequestConsts.WOPI_CONTENTS_PATH))
                {
                    // The rawId ends with /contents so this is a request to read/write the file contents

                    // Remove /contents from the end of rawId to get the actual file id
                    requestData.Id = rawId.Substring(0, rawId.Length - WopiRequestConsts.WOPI_CONTENTS_PATH.Length);

                    // Check request verb to determine file operation
                    if (request.HttpMethod == "GET")
                        requestData.RequestType = WopiRequestType.GetFile;
                    if (request.HttpMethod == "POST")
                        requestData.RequestType = WopiRequestType.PutFile;
                }
                else
                {
                    requestData.Id = rawId;

                    if (request.HttpMethod == "GET")
                    {
                        // GET requests to the file are always CheckFileInfo
                        requestData.RequestType = WopiRequestType.CheckFileInfo;
                    }
                    else if (request.HttpMethod == "POST")
                    {
                        // Use the X-WOPI-Override header to determine the request type for POSTs
                        string wopiOverride = request.Headers[WopiRequestHeaders.OVERRIDE];
                        switch (wopiOverride)
                        {
                            case "LOCK":
                                // Check lock type based on presence of OldLock header
                                if (request.Headers[WopiRequestHeaders.OLD_LOCK] != null)
                                    requestData.RequestType = WopiRequestType.UnlockAndRelock;
                                else
                                    requestData.RequestType = WopiRequestType.Lock;
                                break;
                            case "GET_LOCK":
                                requestData.RequestType = WopiRequestType.GetLock;
                                break;
                            case "REFRESH_LOCK":
                                requestData.RequestType = WopiRequestType.RefreshLock;
                                break;
                            case "UNLOCK":
                                requestData.RequestType = WopiRequestType.Unlock;
                                break;
                            case "PUT_RELATIVE":
                                requestData.RequestType = WopiRequestType.PutRelativeFile;
                                break;
                            case "RENAME_FILE":
                                requestData.RequestType = WopiRequestType.RenameFile;
                                break;
                            case "PUT_USER_INFO":
                                requestData.RequestType = WopiRequestType.PutUserInfo;
                                break;

                                /*
                                // The following WOPI_Override values were referenced in the product group sample, but not in the documentation
                                case "COBALT":
                                    requestData.RequestType = WopiRequestType.ExecuteCobaltRequest;
                                    break;
                                case "DELETE":
                                    requestData.RequestType = WopiRequestType.DeleteFile;
                                    break;
                                case "READ_SECURE_STORE":
                                    requestData.RequestType = WopiRequestType.ReadSecureStore;
                                    break;
                                case "GET_RESTRICTED_LINK":
                                    requestData.RequestType = WopiRequestType.GetRestrictedLink;
                                    break;
                                case "REVOKE_RESTRICTED_LINK":
                                    requestData.RequestType = WopiRequestType.RevokeRestrictedLink;
                                    break;
                                */
                        }
                    }
                }
            }
            else if (wopiPath.StartsWith(WopiRequestConsts.WOPI_FOLDERS_PATH))
            {
                // This is a folder-related request

                // Remove /folders/ from the beginning of wopiPath
                string rawId = wopiPath.Substring(WopiRequestConsts.WOPI_FOLDERS_PATH.Length);

                if (rawId.EndsWith(WopiRequestConsts.WOPI_CHILDREN_PATH))
                {
                    // rawId ends with /children, so it's an EnumerateChildren request.

                    // Remove /children from the end of rawId
                    requestData.Id = rawId.Substring(0, WopiRequestConsts.WOPI_CHILDREN_PATH.Length);
                    //requestData.RequestType = WopiRequestType.EnumerateChildren;
                }
                else
                {
                    // rawId doesn't end with /children, so it's a CheckFolderInfo.

                    requestData.Id = rawId;
                    //requestData.RequestType = WopiRequestType.CheckFolderInfo;
                }
            }
            else
            {
                // This is an unknown request
                requestData.RequestType = WopiRequestType.None;
            }

            return requestData;
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