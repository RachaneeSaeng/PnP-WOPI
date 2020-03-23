using com.microsoft.dx.officewopi.Models.Wopi;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Runtime.Caching;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Xml.Linq;

namespace com.microsoft.dx.officewopi.Utils
{
    /// <summary>
    /// Contains methods to support request handleing from Wopi
    /// </summary>
    public static class WopiRequestUtil
    {


        /// <summary>
        /// Validates the WOPI Proof on an incoming WOPI request
        /// </summary>
        public async static Task<bool> ValidateWopiProof(HttpContext context)
        {
            // Make sure the request has the correct headers
            if (context.Request.Headers[WopiRequestHeaders.PROOF] == null ||
                context.Request.Headers[WopiRequestHeaders.TIME_STAMP] == null)
                return false;

            // Set the requested proof values
            var requestProof = context.Request.Headers[WopiRequestHeaders.PROOF];
            var requestProofOld = String.Empty;
            if (context.Request.Headers[WopiRequestHeaders.PROOF_OLD] != null)
                requestProofOld = context.Request.Headers[WopiRequestHeaders.PROOF_OLD];

            // Get the WOPI proof info from discovery
            var discoProofPublicKey = await getWopiProofPublicKey();

            // Encode the values into bytes
            var accessTokenBytes = Encoding.UTF8.GetBytes(context.Request.QueryString["access_token"]);

            var hostUrl = context.Request.Url.OriginalString.Replace(":44300", "").Replace(":443", "");
            var hostUrlBytes = Encoding.UTF8.GetBytes(hostUrl.ToUpperInvariant());

            var timeStampBytes = BitConverter.GetBytes(Convert.ToInt64(context.Request.Headers[WopiRequestHeaders.TIME_STAMP])).Reverse().ToArray();

            // Build expected proof
            List<byte> expected = new List<byte>(
                4 + accessTokenBytes.Length +
                4 + hostUrlBytes.Length +
                4 + timeStampBytes.Length);

            // Add the values to the expected variable
            expected.AddRange(BitConverter.GetBytes(accessTokenBytes.Length).Reverse().ToArray());
            expected.AddRange(accessTokenBytes);
            expected.AddRange(BitConverter.GetBytes(hostUrlBytes.Length).Reverse().ToArray());
            expected.AddRange(hostUrlBytes);
            expected.AddRange(BitConverter.GetBytes(timeStampBytes.Length).Reverse().ToArray());
            expected.AddRange(timeStampBytes);
            byte[] expectedBytes = expected.ToArray();

            return (verifyProof(expectedBytes, requestProof, discoProofPublicKey.value) ||
                verifyProof(expectedBytes, requestProofOld, discoProofPublicKey.value) ||
                verifyProof(expectedBytes, requestProof, discoProofPublicKey.oldvalue));
        }


        /// <summary>
        /// Gets the WOPI proof details from the WOPI discovery endpoint and caches it appropriately
        /// </summary>
        private async static Task<WopiProofPublicKey> getWopiProofPublicKey()
        {
            // Check cache for this data
            MemoryCache memoryCache = MemoryCache.Default;
            if (memoryCache.Contains("WopiProof"))
                return (WopiProofPublicKey)memoryCache["WopiProof"];

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
                    var proof = discoXml.Descendants("proof-key").FirstOrDefault();
                    var wopiProof = new WopiProofPublicKey()
                    {
                        value = proof.Attribute("value").Value,
                        //modulus = proof.Attribute("modulus").Value,
                        //exponent = proof.Attribute("exponent").Value,
                        oldvalue = proof.Attribute("oldvalue").Value,
                        //oldmodulus = proof.Attribute("oldmodulus").Value,
                        //oldexponent = proof.Attribute("oldexponent").Value
                    };

                    // Add to cache for 20min
                    memoryCache.Add("WopiProof", wopiProof, DateTimeOffset.Now.AddMinutes(20));

                    return wopiProof;
                }
            }

            return null;

        }

        /// <summary>
        /// Verifies the proof against a specified key
        /// </summary>
        private static bool verifyProof(byte[] expectedProof, string proofFromRequest, string discoPublicKey)
        {
            using (RSACryptoServiceProvider rsaProvider = new RSACryptoServiceProvider())
            {
                try
                {
                    rsaProvider.ImportCspBlob(Convert.FromBase64String(discoPublicKey));
                    return rsaProvider.VerifyData(expectedProof, "SHA256", Convert.FromBase64String(proofFromRequest));
                }
                catch (FormatException)
                {
                    return false;
                }
                catch (CryptographicException)
                {
                    return false;
                }
            }
        }

    }

}