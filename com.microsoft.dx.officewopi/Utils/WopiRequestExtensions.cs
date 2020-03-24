using com.microsoft.dx.officewopi.Models;
using com.microsoft.dx.officewopi.Security;
using Newtonsoft.Json;
using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;

namespace com.microsoft.dx.officewopi.Utils
{
    /// <summary>
    /// Provides processing extensions for each of the WOPI operations
    /// </summary>
    public static class WopiRequestExtensions
    {
        /// <summary>
        /// The CheckFileInfo operation is one of the most important WOPI operations. 
        /// This operation must be implemented for all WOPI actions. 
        /// CheckFileInfo returns information about a file, a user’s permissions on that file, 
        /// and general information about the capabilities that the WOPI host has on the file. 
        /// In addition, some CheckFileInfo properties can influence the appearance and behavior of WOPI clients.
        /// </summary>
        /// <remarks>
        /// For full documentation on CheckFileInfo, see https://wopi.readthedocs.org/projects/wopirest/en/latest/files/CheckFileInfo.html
        /// </remarks>
        public static HttpResponseMessage CheckFileInfo(this HttpContext context, FileModel file)
        {
            // Serialize the response object
            string jsonString = JsonConvert.SerializeObject(file);

            // Write the response and return a success 200
            var response = ReturnStatus(HttpStatusCode.OK, "Success");
            response.Content = new StringContent(jsonString);
            return response;
        }

        /// <summary>
        /// The GetFile operation retrieves a file from a host.
        /// </summary>
        /// <remarks>
        /// For full documentation on GetFile, see https://wopi.readthedocs.org/projects/wopirest/en/latest/files/GetFile.html
        /// </remarks>
        public async static Task<HttpResponseMessage> GetFile(this HttpContext context, FileModel file)
        {
            // Get the file from blob storage
            var bytes = await AzureStorageUtil.GetFile(file.id.ToString(), file.Container);

            // Write the response and return success 200
            var response = ReturnStatus(HttpStatusCode.OK, "Success");
            response.Content = new ByteArrayContent(bytes);
            return response;
        }

        /// <summary>
        /// The Lock operation locks a file for editing by the WOPI client application instance that requested the lock. 
        /// To support editing files, WOPI clients require that the WOPI host support locking files. 
        /// When locked, a file should not be writable by other applications.
        /// </summary>
        /// <remarks>
        /// For full documentation on Lock, see https://wopi.readthedocs.org/projects/wopirest/en/latest/files/Lock.html
        /// </remarks>
        public async static Task<HttpResponseMessage> Lock(this HttpContext context, FileModel file)
        {
            // Get the Lock value passed in on the request
            string requestLock = context.Request.Headers[WopiRequestHeaders.LOCK];

            // Ensure the file isn't already locked or expired
            if (String.IsNullOrEmpty(file.LockValue) ||
                (file.LockExpires != null &&
                file.LockExpires < DateTime.Now))
            {
                // Update the file with a LockValue and LockExpiration
                file.LockValue = requestLock;
                file.LockExpires = DateTime.Now.AddMinutes(30);
                await DocumentDBRepository<FileModel>.UpdateItemAsync("Files", file.id.ToString(), (FileModel)file);

                // Return success 200
                return ReturnStatus(HttpStatusCode.OK, "Success");
            }
            else if (file.LockValue == requestLock)
            {
                // File lock matches existing lock, so refresh lock by extending expiration
                file.LockExpires = DateTime.Now.AddMinutes(30);
                await DocumentDBRepository<FileModel>.UpdateItemAsync("Files", file.id.ToString(), (FileModel)file);

                // Return success 200
                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                return ReturnStatus(HttpStatusCode.OK, "Success");
            }
            else
            {
                // The file is locked by someone else...return mismatch
                return ReturnLockMismatch(file.LockValue, String.Format("File already locked by {0}", file.LockValue));
            }
        }

        /// <summary>
        /// The GetLock operation retrieves a lock on a file. Note that this operation does not create a new lock. 
        /// Rather, this operation always returns the current lock value in the X-WOPI-Lock response header.
        /// </summary>
        /// <remarks>
        /// For full documentation on GetLock, see https://wopi.readthedocs.org/projects/wopirest/en/latest/files/GetLock.html
        /// </remarks>
        public async static Task<HttpResponseMessage> GetLock(this HttpContext context, FileModel file)
        {
            // Check for valid lock on file
            if (String.IsNullOrEmpty(file.LockValue))
            {
                // File is not locked...return empty X-WOPI-Lock header
                context.Response.Headers[WopiResponseHeaders.LOCK] = String.Empty;

                // Return success 200
                return ReturnStatus(HttpStatusCode.OK, "Success");
            }
            else if (file.LockExpires != null && file.LockExpires < DateTime.Now)
            {
                // File lock expired, so clear it out
                file.LockValue = null;
                file.LockExpires = null;
                await DocumentDBRepository<FileModel>.UpdateItemAsync("Files", file.id.ToString(), (FileModel)file);

                // File is not locked...return empty X-WOPI-Lock header
                context.Response.Headers[WopiResponseHeaders.LOCK] = String.Empty;

                // Return success 200
                return ReturnStatus(HttpStatusCode.OK, "Success");
            }
            else
            {
                // File has a valid lock, so we need to return it
                context.Response.Headers[WopiResponseHeaders.LOCK] = file.LockValue;

                // Return success 200
                return ReturnStatus(HttpStatusCode.OK, "Success");
            }
        }

        /// <summary>
        /// The RefreshLock operation refreshes the lock on a file by resetting its automatic expiration timer to 30 minutes.
        /// </summary>
        /// <remarks>
        /// For full documentation on RefreshLock, see https://wopi.readthedocs.org/projects/wopirest/en/latest/files/RefreshLock.html
        /// </remarks>
        public async static Task<HttpResponseMessage> RefreshLock(this HttpContext context, FileModel file)
        {
            // Get the Lock value passed in on the request
            string requestLock = context.Request.Headers[WopiRequestHeaders.LOCK];

            // Ensure the file has a valid lock
            if (String.IsNullOrEmpty(file.LockValue))
            {
                // File isn't locked...pass empty Lock in mismatch response
                return ReturnLockMismatch(String.Empty, "File isn't locked");
            }
            else if (file.LockExpires != null && file.LockExpires < DateTime.Now)
            {
                // File lock expired, so clear it out
                file.LockValue = null;
                file.LockExpires = null;
                await DocumentDBRepository<FileModel>.UpdateItemAsync("Files", file.id.ToString(), (FileModel)file);

                // File isn't locked...pass empty Lock in mismatch response
                return ReturnLockMismatch(String.Empty, "File isn't locked");
            }
            else if (requestLock != file.LockValue)
            {
                // File lock mismatch...pass Lock in mismatch response
                return ReturnLockMismatch(file.LockValue, "Lock mismatch");
            }
            else
            {
                // Extend the expiration
                file.LockExpires = DateTime.Now.AddMinutes(30);
                await DocumentDBRepository<FileModel>.UpdateItemAsync("Files", file.id.ToString(), (FileModel)file);

                // Return success 200
                return ReturnStatus(HttpStatusCode.OK, "Success");
            }
        }

        /// <summary>
        /// The Unlock operation releases the lock on a file.
        /// </summary>
        /// <remarks>
        /// For full documentation on Unlock, see https://wopi.readthedocs.org/projects/wopirest/en/latest/files/Unlock.html
        /// </remarks>
        public async static Task<HttpResponseMessage> Unlock(this HttpContext context, FileModel file)
        {
            // Get the Lock value passed in on the request
            string requestLock = context.Request.Headers[WopiRequestHeaders.LOCK];

            // Ensure the file has a valid lock
            if (String.IsNullOrEmpty(file.LockValue))
            {
                // File isn't locked...pass empty Lock in mismatch response
                return ReturnLockMismatch(String.Empty, "File isn't locked");
            }
            else if (file.LockExpires != null && file.LockExpires < DateTime.Now)
            {
                // File lock expired, so clear it out
                file.LockValue = null;
                file.LockExpires = null;
                await DocumentDBRepository<FileModel>.UpdateItemAsync("Files", file.id.ToString(), (FileModel)file);

                // File isn't locked...pass empty Lock in mismatch response
                return ReturnLockMismatch(String.Empty, "File isn't locked");
            }
            else if (requestLock != file.LockValue)
            {
                // File lock mismatch...pass Lock in mismatch response
                return ReturnLockMismatch(file.LockValue, "Lock mismatch");
            }
            else
            {
                // Unlock the file
                file.LockValue = null;
                file.LockExpires = null;
                await DocumentDBRepository<FileModel>.UpdateItemAsync("Files", file.id.ToString(), (FileModel)file);

                // Return success 200
                return ReturnStatus(HttpStatusCode.OK, "Success");
            }
        }

        /// <summary>
        /// The UnlockAndRelock operation releases a lock on a file, and then immediately takes a new lock on the file.
        /// </summary>
        /// <remarks>
        /// For full documentation on UnlockAndRelock, see https://wopi.readthedocs.org/projects/wopirest/en/latest/files/UnlockAndRelock.html
        /// </remarks>
        public async static Task<HttpResponseMessage> UnlockAndRelock(this HttpContext context, FileModel file)
        {
            // Get the Lock and OldLock values passed in on the request
            string requestLock = context.Request.Headers[WopiRequestHeaders.LOCK];
            string requestOldLock = context.Request.Headers[WopiRequestHeaders.OLD_LOCK];

            // Ensure the file has a valid lock
            if (String.IsNullOrEmpty(file.LockValue))
            {
                // File isn't locked...pass empty Lock in mismatch response
                return ReturnLockMismatch(String.Empty, "File isn't locked");
            }
            else if (file.LockExpires != null && file.LockExpires < DateTime.Now)
            {
                // File lock expired, so clear it out
                file.LockValue = null;
                file.LockExpires = null;
                await DocumentDBRepository<FileModel>.UpdateItemAsync("Files", file.id.ToString(), (FileModel)file);

                // File isn't locked...pass empty Lock in mismatch response
                return ReturnLockMismatch(String.Empty, "File isn't locked");
            }
            else if (requestOldLock != file.LockValue)
            {
                // File lock mismatch...pass Lock in mismatch response
                return ReturnLockMismatch(file.LockValue, "Lock mismatch");
            }
            else
            {
                // Update the file with a LockValue and LockExpiration
                file.LockValue = requestLock;
                file.LockExpires = DateTime.Now.AddMinutes(30);
                await DocumentDBRepository<FileModel>.UpdateItemAsync("Files", file.id.ToString(), (FileModel)file);

                // Return success 200
                return ReturnStatus(HttpStatusCode.OK, "Success");
            }
        }

        /// <summary>
        /// The PutFile operation updates a file’s binary contents.
        /// </summary>
        /// <remarks>
        /// For full documentation on PutFile, see https://wopi.readthedocs.org/projects/wopirest/en/latest/files/PutFile.html
        /// </remarks>
        public async static Task<HttpResponseMessage> PutFile(this HttpContext context, FileModel file)
        {
            // Get the Lock value passed in on the request
            string requestLock = context.Request.Headers[WopiRequestHeaders.LOCK];
            var editors = context.Request.Headers[WopiRequestHeaders.EDITORS];

            // Ensure the file has a valid lock
            if (String.IsNullOrEmpty(file.LockValue))
            {
                // If the file is 0 bytes, this is document creation
                if (file.Size == 0)
                {
                    // Update the file in blob storage
                    var bytes = new byte[context.Request.InputStream.Length];
                    context.Request.InputStream.Read(bytes, 0, bytes.Length);
                    file.Size = bytes.Length;
                    await AzureStorageUtil.UploadFile(file.id.ToString(), file.Container, bytes);

                    // Update version
                    file.Version++;
                    file.LastModifiedTime = DateTime.Now;
                    file.LastModifiedUsers = editors;
                    await DocumentDBRepository<FileModel>.UpdateItemAsync("Files", file.id.ToString(), (FileModel)file);

                    // Return success 200
                    return ReturnStatus(HttpStatusCode.OK, "Success");
                }
                else
                {
                    // File isn't locked...pass empty Lock in mismatch response
                    return ReturnLockMismatch(String.Empty, "File isn't locked");
                }
            }
            else if (file.LockExpires != null && file.LockExpires < DateTime.Now)
            {
                // File lock expired, so clear it out
                file.LockValue = null;
                file.LockExpires = null;
                await DocumentDBRepository<FileModel>.UpdateItemAsync("Files", file.id.ToString(), (FileModel)file);

                // File isn't locked...pass empty Lock in mismatch response
                return ReturnLockMismatch(String.Empty, "File isn't locked");
            }
            else if (requestLock != file.LockValue)
            {
                // File lock mismatch...pass Lock in mismatch response
                return ReturnLockMismatch(file.LockValue, "Lock mismatch");
            }
            else
            {
                // Update the file in blob storage
                var bytes = new byte[context.Request.InputStream.Length];
                context.Request.InputStream.Read(bytes, 0, bytes.Length);
                file.Size = bytes.Length;
                await AzureStorageUtil.UploadFile(file.id.ToString(), file.Container, bytes);

                // Update version
                file.Version++;
                file.LastModifiedTime = DateTime.Now;
                file.LastModifiedUsers = editors;
                await DocumentDBRepository<FileModel>.UpdateItemAsync("Files", file.id.ToString(), (FileModel)file);

                // Return success 200
                return ReturnStatus(HttpStatusCode.OK, "Success");
            }
        }

        /// <summary>
        /// The PutRelativeFile operation creates a new file on the host based on the current file
        /// </summary>
        /// <remarks>
        /// For full documentation on PutRelativeFile, see https://wopi.readthedocs.org/projects/wopirest/en/latest/files/PutRelativeFile.html
        /// </remarks>
        public async static Task<HttpResponseMessage> PutRelativeFile(this HttpContext context, DetailedFileModel file)
        {
            // Determine the specific mode
            if (context.Request.Headers[WopiRequestHeaders.RELATIVE_TARGET] != null &&
                context.Request.Headers[WopiRequestHeaders.SUGGESTED_TARGET] != null)
            {
                // Theses headers are mutually exclusive, so we should return a 501 Not Implemented
                return ReturnStatus(HttpStatusCode.NotImplemented, "Both RELATIVE_TARGET and SUGGESTED_TARGET were present");
            }
            else if (context.Request.Headers[WopiRequestHeaders.RELATIVE_TARGET] != null ||
                context.Request.Headers[WopiRequestHeaders.SUGGESTED_TARGET] != null)
            {
                string fileName = "";
                if (context.Request.Headers[WopiRequestHeaders.RELATIVE_TARGET] != null)
                {
                    // Specific mode...use the exact filename
                    fileName = context.Request.Headers[WopiRequestHeaders.RELATIVE_TARGET];
                }
                else
                {
                    // Suggested mode...might just be an extension
                    fileName = context.Request.Headers[WopiRequestHeaders.RELATIVE_TARGET];
                    if (fileName.IndexOf('.') == 0)
                        fileName = file.BaseFileName.Substring(0, file.BaseFileName.LastIndexOf('.')) + fileName;
                }

                // Create the file entity
                DetailedFileModel newFile = new DetailedFileModel()
                {
                    id = Guid.NewGuid(),
                    OwnerId = file.OwnerId,
                    BaseFileName = fileName,
                    Size = context.Request.InputStream.Length,
                    Container = file.Container,
                    Version = 1
                };

                // First stream the file into blob storage
                var stream = context.Request.InputStream;
                var bytes = new byte[stream.Length];
                await stream.ReadAsync(bytes, 0, (int)stream.Length);
                var id = await Utils.AzureStorageUtil.UploadFile(newFile.id.ToString(), newFile.Container, bytes);

                // Write the details into documentDB
                await DocumentDBRepository<FileModel>.CreateItemAsync("Files", (FileModel)newFile);

                // Get access token for the new file
                WopiSecurity security = new WopiSecurity();
                var token = security.GenerateToken(newFile.OwnerId, newFile.Container, newFile.id.ToString());
                var tokenStr = security.WriteToken(token);

                var view = file.Actions.FirstOrDefault(i => i.name == "view");
                var edit = file.Actions.FirstOrDefault(i => i.name == "edit");

                // Prepare the Json response
                var jsonObj = new
                {
                    Name = newFile.BaseFileName,
                    Url = $@"https://{context.Request.Url.Authority}/wopi/files/{newFile.id}?access_token={tokenStr}",
                    HostViewUrl = view != null ? $@"https://{context.Request.Url.Authority}/Home/Detail/{newFile.id}?action=view" : "",
                    HostEditUrl = edit != null ? $@"https://{context.Request.Url.Authority}/Home/Detail/{newFile.id}?action=edit" : ""
                };


                // Write the response and return a success 200
                var response = ReturnStatus(HttpStatusCode.OK, "Success");
                response.Content = new StringContent(JsonConvert.SerializeObject(jsonObj));
                return response;
            }
            else
            {
                return ReturnStatus(HttpStatusCode.BadRequest, "PutRelativeFile mode was not provided in the request");
            }
        }

        /// <summary>
        /// The RenameFile operation renames a file.
        /// </summary>
        /// <remarks>
        /// For full documentation on RenameFile, see https://wopi.readthedocs.org/projects/wopirest/en/latest/files/RenameFile.html
        /// </remarks>
        public async static Task<HttpResponseMessage> RenameFile(this HttpContext context, FileModel file)
        {
            // Get the Lock value passed in on the request
            string requestLock = context.Request.Headers[WopiRequestHeaders.LOCK];

            // Make sure the X-WOPI-RequestedName header is included
            if (context.Request.Headers[WopiRequestHeaders.REQUESTED_NAME] != null)
            {
                // Get the new file name
                var newFileName = context.Request.Headers[WopiRequestHeaders.REQUESTED_NAME];
                var fileExt = file.BaseFileName.Substring(file.BaseFileName.LastIndexOf('.') + 1).ToLower();

                // Ensure the file isn't locked
                if (String.IsNullOrEmpty(file.LockValue) ||
                    (file.LockExpires != null &&
                    file.LockExpires < DateTime.Now))
                {
                    // Update the file with a LockValue and LockExpiration
                    file.LockValue = requestLock;
                    file.LockExpires = DateTime.Now.AddMinutes(30);
                    file.BaseFileName = $"{newFileName}.{fileExt}";
                    await DocumentDBRepository<FileModel>.UpdateItemAsync("Files", file.id.ToString(), (FileModel)file);

                    // Return success 200
                    var jsonObj = new { Name = newFileName };

                    var response = ReturnStatus(HttpStatusCode.OK, "Success");
                    response.Content = new StringContent(JsonConvert.SerializeObject(jsonObj));
                    return response;
                }
                else if (file.LockValue == requestLock)
                {
                    // File lock matches existing lock, so we can change the name
                    file.LockExpires = DateTime.Now.AddMinutes(30);
                    file.BaseFileName = $"{newFileName}.{fileExt}";
                    await DocumentDBRepository<FileModel>.UpdateItemAsync("Files", file.id.ToString(), (FileModel)file);

                    // Return success 200
                    var jsonObj = new { Name = newFileName };

                    var response = ReturnStatus(HttpStatusCode.OK, "Success");
                    response.Content = new StringContent(JsonConvert.SerializeObject(jsonObj));
                    return response;
                }
                else
                {
                    // The file is locked by someone else...return mismatch
                    return ReturnLockMismatch(file.LockValue, String.Format("File locked by {0}", file.LockValue));
                }
            }
            else
            {
                // X-WOPI-RequestedName header wasn't included
                return ReturnStatus(HttpStatusCode.BadRequest, "X-WOPI-RequestedName header wasn't included in request");
            }
        }

        /// <summary>
        /// The PutUserInfo operation stores some basic user information on the host. 
        /// When a host receives this request, they must store the UserInfo string which is contained in the body of the request. 
        /// The UserInfo string should be associated with a particular user, and should be passed back to the WOPI client in subsequent CheckFileInfo responses in the UserInfo property.
        /// </summary>
        /// <remarks>
        /// For full documentation on PutUserInfo, see https://wopi.readthedocs.org/projects/wopirest/en/latest/files/PutUserInfo.html
        /// </remarks>
        public async static Task<HttpResponseMessage> PutUserInfo(this HttpContext context, FileModel file)
        {
            // Set and save the UserInfo on the file
            var stream = context.Request.InputStream;
            var bytes = new byte[stream.Length];
            await stream.ReadAsync(bytes, 0, (int)stream.Length);
            //file.UserInfo = System.Text.Encoding.UTF8.GetString(bytes);

            // Update the file in DocumentDB
            await DocumentDBRepository<FileModel>.UpdateItemAsync("Files", file.id.ToString(), (FileModel)file);

            // Return success
            return ReturnStatus(HttpStatusCode.OK, "Success");
        }


        /// <summary>
        /// Handles mismatch responses on WOPI requests
        /// </summary>
        private static HttpResponseMessage ReturnLockMismatch(string existingLock = null, string reason = null)
        {
            var response = ReturnStatus(HttpStatusCode.Conflict, "Lock mismatch/Locked by another interface");
            response.Headers.Add(WopiResponseHeaders.LOCK, existingLock ?? String.Empty);
            if (!String.IsNullOrEmpty(reason))
            {
                response.Headers.Add(WopiResponseHeaders.LOCK_FAILURE_REASON, reason);
            }
            return response;
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
