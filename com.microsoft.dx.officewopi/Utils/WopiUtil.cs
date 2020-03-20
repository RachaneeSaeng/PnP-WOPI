﻿using com.microsoft.dx.officewopi.Models;
using com.microsoft.dx.officewopi.Models.Wopi;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Runtime.Caching;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace com.microsoft.dx.officewopi.Utils
{
    /// <summary>
    /// Contains method to support generating iframe request to Wopi
    /// </summary>
    public static class WopiUtil
    {
        /// <summary>
        /// Populates a list of files with action details from WOPI discovery
        /// </summary>
        public async static Task PopulateActions(this IEnumerable<DetailedFileModel> files)
        {
            if (files.Count() > 0)
            {
                foreach (var file in files)
                {
                    await file.PopulateActions();
                }
            }
        }

        /// <summary>
        /// Populates a file with action details from WOPI discovery based on the file extension
        /// </summary>
        public async static Task PopulateActions(this DetailedFileModel file)
        {
            // Get the discovery informations
            var actions = await GetDiscoveryInfo();
            var fileExt = file.BaseFileName.Substring(file.BaseFileName.LastIndexOf('.') + 1).ToLower();
            file.Actions = actions.Where(i => i.ext == fileExt).OrderBy(i => i.isDefault).ToList();
        }

        /// <summary>
        /// Gets the discovery information from WOPI discovery and caches it appropriately
        /// </summary>
        private async static Task<List<WopiAction>> GetDiscoveryInfo()
        {
            List<WopiAction> actions = new List<WopiAction>();

            // Determine if the discovery data is cached
            MemoryCache memoryCache = MemoryCache.Default;
            if (memoryCache.Contains("DiscoData"))
                actions = (List<WopiAction>)memoryCache["DiscoData"];
            else
            {
                // Data isn't cached, so we will use the Wopi Discovery endpoint to get the data
                HttpClient client = new HttpClient();
                using (HttpResponseMessage response = await client.GetAsync(ConfigurationManager.AppSettings["WopiDiscovery"]))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        // Read the xml string from the response
                        string xmlString = await response.Content.ReadAsStringAsync();

                        // Parse the xml string into Xml
                        var discoXml = XDocument.Parse(xmlString);

                        // Convert the discovery xml into list of WopiApp
                        var xapps = discoXml.Descendants("app");
                        foreach (var xapp in xapps)
                        {
                            // Parse the actions for the app
                            var xactions = xapp.Descendants("action");
                            foreach (var xaction in xactions)
                            {
                                actions.Add(new WopiAction()
                                {
                                    app = xapp.Attribute("name").Value,
                                    favIconUrl = xapp.Attribute("favIconUrl").Value,
                                    checkLicense = Convert.ToBoolean(xapp.Attribute("checkLicense").Value),
                                    name = xaction.Attribute("name").Value,
                                    ext = (xaction.Attribute("ext") != null) ? xaction.Attribute("ext").Value : String.Empty,
                                    progid = (xaction.Attribute("progid") != null) ? xaction.Attribute("progid").Value : String.Empty,
                                    isDefault = (xaction.Attribute("default") != null) ? true : false,
                                    urlsrc = xaction.Attribute("urlsrc").Value,
                                    requires = (xaction.Attribute("requires") != null) ? xaction.Attribute("requires").Value : String.Empty
                                });
                            }

                            // Cache the discovey data for an hour
                            memoryCache.Add("DiscoData", actions, DateTimeOffset.Now.AddHours(1));
                        }
                    }
                }
            }

            return actions;
        }

        /// <summary>
        /// Forms the correct action url for the file and host
        /// </summary>
        public static string GetActionUrl(WopiAction action, FileModel file, string authority)
        {
            // Initialize the urlsrc
            var urlsrc = action.urlsrc;

            // Look through the action placeholders
            var phCnt = 0;
            foreach (var p in WopiUrlPlaceholders.Placeholders)
            {
                if (urlsrc.Contains(p))
                {
                    // Replace the placeholder value accordingly
                    var ph = GetPlaceholderValue(p, file, authority);
                    if (!String.IsNullOrEmpty(ph))
                    {
                        urlsrc = urlsrc.Replace(p, ph + "&");
                        phCnt++;
                    }
                    else
                        urlsrc = urlsrc.Replace(p, ph);
                }
            }

            // Add the WOPISrc to the end of the request
            //urlsrc += ((phCnt > 0) ? "" : "?"); //+ String.Format("WOPISrc=https://{0}/wopi/files/{1}", authority, file.id.ToString());
            return urlsrc;
        }

        /// <summary>
        /// Sets a specific WOPI URL placeholder with the correct value
        /// Most of these are hard-coded in this WOPI implementation
        /// </summary>
        private static string GetPlaceholderValue(string placeholder, FileModel file, string authority)
        {
            var ph = placeholder.Substring(1, placeholder.IndexOf("="));

            switch (placeholder)
            {
                case WopiUrlPlaceholders.BUSINESS_USER:
                    return ph + "1";
                case WopiUrlPlaceholders.DC_LLCC:
                case WopiUrlPlaceholders.UI_LLCC:
                    return ph + "en-US";
                case WopiUrlPlaceholders.THEME_ID:
                    return ph + "1";
                case WopiUrlPlaceholders.DISABLE_CHAT:
                    return ph + "0";
                case WopiUrlPlaceholders.PERFSTATS:
                    return ph + "0";
                case WopiUrlPlaceholders.VALIDATOR_TEST_CATEGORY:
                    return ph + "OfficeOnline"; //This value can be set to All, OfficeOnline or OfficeNativeClient to activate tests specific to Office Online and Office for iOS. If omitted, the default value is All.                   
                case WopiUrlPlaceholders.WOPI_SOURCE:
                    return String.Format("{0}https://{1}/wopi/files/{2}", ph, authority, file.id.ToString());
                default:
                    return "";

            }
        }
    }
}
