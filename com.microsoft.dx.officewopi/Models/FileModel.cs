using Newtonsoft.Json;
using System;

namespace com.microsoft.dx.officewopi.Models
{
    /// <summary>
    /// This class contains the file properties that are persisted in DocumentDB
    /// </summary>
    public class FileModel
    {
        [JsonProperty(PropertyName = "id")]
        public Guid id { get; set; }

        [JsonProperty(PropertyName = "LockValue")]
        public string LockValue { get; set; }

        [JsonProperty(PropertyName = "LockExpires")]
        public DateTime? LockExpires { get; set; }

        [JsonProperty(PropertyName = "OwnerId")]
        public string OwnerId { get; set; }

        [JsonProperty(PropertyName = "BaseFileName")]
        public string BaseFileName { get; set; }

        [JsonProperty(PropertyName = "Container")]
        public string Container { get; set; }

        [JsonProperty(PropertyName = "Size")]
        public long Size { get; set; }

        [JsonProperty(PropertyName = "Version")]
        public int Version { get; set; }

        [JsonProperty(PropertyName = "LastModifiedTime")]
        public DateTime LastModifiedTime { get; set; }

        [JsonProperty(PropertyName = "LastModifiedUsers")]
        public string LastModifiedUsers { get; set; }

        //[JsonProperty(PropertyName = "UserInfo")]
        //public string UserInfo { get; set; }

    }
}
