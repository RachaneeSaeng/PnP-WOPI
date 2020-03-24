using com.microsoft.dx.officewopi.Models.Wopi;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace com.microsoft.dx.officewopi.Models
{
    /// <summary>
    /// This class contains additional file properties that are used in 
    /// CheckFileInfo requests, but not persisted in the database.
    /// 
    /// Note: many of the properties are hard-coded for this WOPI Host and
    /// might be more dynamic in other implementations
    /// </summary>
    public class DetailedFileModel : FileModel
    {
        [JsonProperty(PropertyName = "UserId")]
        public string UserId { get; set; }

        [JsonProperty(PropertyName = "CloseUrl")]
        public string CloseUrl { get; set; }

        [JsonProperty(PropertyName = "HostEditUrl")]
        public string HostEditUrl { get; set; }

        [JsonProperty(PropertyName = "HostViewUrl")]
        public string HostViewUrl { get; set; }

        [JsonProperty(PropertyName = "HostEmbeddedViewUrl")]
        public string HostEmbeddedViewUrl { get; set; }

        [JsonProperty(PropertyName = "FileVersionUrl")]
        public string FileVersionUrl { get; set; }

        /// <summary>
        /// the DownloadUrl is used to provide a direct download link to the file if the user’s subscription check fails.
        /// </summary>
        [JsonProperty(PropertyName = "DownloadUrl")]
        public string DownloadUrl { get; set; }

        [JsonProperty(PropertyName = "SupportsCoauth")]
        public bool SupportsCoauth
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "SupportsExtendedLockLength")]
        public bool SupportsExtendedLockLength
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "SupportsFileCreation")]
        public bool SupportsFileCreation
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "SupportsFolders")]
        public bool SupportsFolders
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "SupportsGetLock")]
        public bool SupportsGetLock
        {
            get { return true; }
        }

        [JsonProperty(PropertyName = "SupportsLocks")]
        public bool SupportsLocks
        {
            get { return true; }
        }

        [JsonProperty(PropertyName = "SupportsRename")]
        public bool SupportsRename
        {
            get { return true; }
        }

        [JsonProperty(PropertyName = "SupportsScenarioLinks")]
        public bool SupportsScenarioLinks
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "SupportsSecureStore")]
        public bool SupportsSecureStore
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "SupportsUpdate")]
        public bool SupportsUpdate
        {
            get { return true; }
        }

        [JsonProperty(PropertyName = "SupportsUserInfo")]
        public bool SupportsUserInfo
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "LicenseCheckForEditIsEnabled")]
        public bool LicenseCheckForEditIsEnabled
        {
            get { return true; }
        }

        /// <summary>
        /// Permissions for documents
        /// </summary>
        [JsonProperty(PropertyName = "ReadOnly")]
        public bool ReadOnly
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "RestrictedWebViewOnly")]
        public bool RestrictedWebViewOnly
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "UserCanAttend")] //Broadcast only
        public bool UserCanAttend
        {
            get { return true; }
        }

        [JsonProperty(PropertyName = "UserCanNotWriteRelative")]
        public bool UserCanNotWriteRelative
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "UserCanPresent")] //Broadcast only
        public bool UserCanPresent
        {
            get { return true; }
        }

        [JsonProperty(PropertyName = "UserCanRename")]
        public bool UserCanRename
        {
            get { return true; }
        }

        [JsonProperty(PropertyName = "UserCanWrite")]
        public bool UserCanWrite
        {
            get { return true; }
        }

        [JsonProperty(PropertyName = "WebEditingDisabled")]
        public bool WebEditingDisabled
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "Actions")]
        public List<WopiAction> Actions { get; set; }


        [JsonProperty(PropertyName = "UserFriendlyName")]
        public string UserFriendlyName { get; set; }

        [JsonProperty(PropertyName = "BreadcrumbBrandName")]
        public string BreadcrumbBrandName { get; set; }

        [JsonProperty(PropertyName = "BreadcrumbBrandUrl")]
        public string BreadcrumbBrandUrl { get; set; }

        [JsonProperty(PropertyName = "AllowErrorReportPrompt")]
        public bool AllowErrorReportPrompt { get; set; }


    }
}
