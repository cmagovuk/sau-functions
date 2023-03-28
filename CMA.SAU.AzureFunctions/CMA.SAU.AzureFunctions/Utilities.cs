using Azure.Identity;
using Azure.Storage.Blobs;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using PnP.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography.X509Certificates;

namespace CMA.SAU.AzureFunctions
{
    internal class Utilities
    {
        internal static GraphServiceClient GetGraphClientWithCert()
        {
            var scopes = new[] { "https://graph.microsoft.com/.default" };
            string clientId = Environment.GetEnvironmentVariable("CLIENT_ID");
            string tenantId = Environment.GetEnvironmentVariable("TENANT_ID");
            string certThumb = Environment.GetEnvironmentVariable("CERT_THUMBPRINT");


            var clientCertificate = LoadCertificate(StoreName.My, StoreLocation.CurrentUser, certThumb);

            // using Azure.Identity;
            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://docs.microsoft.com/dotnet/api/azure.identity.clientcertificatecredential
            var clientCertCredential = new ClientCertificateCredential(
                tenantId, clientId, clientCertificate, options);

            return new GraphServiceClient(clientCertCredential, scopes);
        }

        internal static X509Certificate2 LoadCertificate(StoreName storeName, StoreLocation storeLocation, string thumbprint)
        {
            // The following code gets the cert from the keystore
            using (X509Store store = new X509Store(storeName, storeLocation))
            {
                store.Open(OpenFlags.ReadOnly);

                X509Certificate2Collection certCollection =
                        store.Certificates.Find(X509FindType.FindByThumbprint,
                        thumbprint, false);

                X509Certificate2Enumerator enumerator = certCollection.GetEnumerator();

                X509Certificate2 cert = null;

                while (enumerator.MoveNext())
                {
                    cert = enumerator.Current;
                }

                return cert;
            }
        }

        internal static string TranslateOne(string term, string translateVar)
        {
            return Translate(term, Translation(translateVar));
        }
        internal static ClientContext GetCaseworkHubContext()
        {
            string sharePointUrl = System.Environment.GetEnvironmentVariable("CASEWORK_HUBSITE");
            return GetContext(sharePointUrl);
        }

        internal static ClientContext GetSAUCasesContext()
        {
            string sharePointUrl = System.Environment.GetEnvironmentVariable("SAU_CASES_SITE");
            return GetContext(sharePointUrl);
        }

        internal static BlobClient GetBlobClient(string blobName)
        {
            string connection = System.Environment.GetEnvironmentVariable("STORAGE_CONN");
            string container = System.Environment.GetEnvironmentVariable("STORAGE_CONTAINER");
            return new BlobClient(connection, container, blobName);
        }

        internal static ClientContext GetContext(string sharePointUrl)
        {
            string clientId = System.Environment.GetEnvironmentVariable("CLIENT_ID");
            string tenantId = System.Environment.GetEnvironmentVariable("TENANT_ID");
            string certThumb = System.Environment.GetEnvironmentVariable("CERT_THUMBPRINT");

            AuthenticationManager auth = new AuthenticationManager(clientId, StoreName.My, StoreLocation.CurrentUser, certThumb, tenantId);

            return auth.GetContext(sharePointUrl);
        }

        internal static void CopyDocumentsToCaseSiteSubFolder(ClientContext caseCtx,ILogger log, dynamic documents, string subFolder)
        {
            foreach (FileDetails file in FormatDocuments(documents))
            {
                CopyDocumentToCaseSite(caseCtx, log, file, subFolder);
            }
        }

        private static void CopyDocumentToCaseSite(ClientContext caseCtx, ILogger log, FileDetails fileDetails, string subFolder)
        {
            try
            {
                BlobClient bc = Utilities.GetBlobClient(fileDetails.Key);
                bc.DownloadContent();

                using (MemoryStream ms = new MemoryStream())
                {
                    bc.DownloadTo(ms);
                    Microsoft.SharePoint.Client.Folder folder = caseCtx.Web.EnsureFolderPath($"Shared Documents/PA Submission/{subFolder}");
                    ms.Position = 0;
                    folder.UploadFile(fileDetails.Filename, ms, true);
                }
            }
            catch (Exception ex)
            {
                log.LogWarning(ex.Message);
                log.LogError("Failed copying document to case site");
            }
        }

        internal static void SendEmail(ClientContext ctx, string emailRecipients, string emailBody, string emailSubject)
        {
            List<string> recipients = new List<string> (emailRecipients.Split(';', StringSplitOptions.RemoveEmptyEntries));
            SendEmail(ctx, recipients, emailBody, emailSubject);
        }

        internal static void SendEmail(ClientContext ctx, List<string> recipients, string emailBody, string emailSubject)
        {
            var ep = new EmailProperties
            {
                To = recipients,
                Subject = emailSubject,
                Body = emailBody
            };

            Utility.SendEmail(ctx, ep);
            ctx.ExecuteQueryRetry();
        }

        private static List<FileDetails> FormatDocuments(dynamic documents)
        {
            List<string> filenames = new List<string>();
            List<FileDetails> fileDetails = new List<FileDetails>();
            if (documents is Newtonsoft.Json.Linq.JArray)
            {
                foreach (dynamic item in documents as Newtonsoft.Json.Linq.JArray)
                {
                    if (item.ContainsKey("key") && item.ContainsKey("filename"))
                    {
                        int index = 1;
                        string filename = GetSafeFilename((string)item.filename);
                        string initFilename = System.IO.Path.GetFileNameWithoutExtension(filename);
                        string extension = System.IO.Path.GetExtension((string)item.filename);
                        while (filenames.Contains(filename))
                        {
                            filename = $"{initFilename} ({index++}){extension}";
                        }
                        filenames.Add(filename);
                        fileDetails.Add(new FileDetails() { Filename = filename, Key = (string)item.key });
                    }
                }
            }
            return fileDetails;
        }

        private static string GetSafeFilename(string filename)
        {
            string safeName = System.IO.Path.GetFileNameWithoutExtension(filename);
            char[] ends = { '.', ' ' };
            safeName = safeName.Trim(ends);

            //Double periods in file name is invalid
            safeName = System.Text.RegularExpressions.Regex.Replace(safeName, @"\.+", ".");
            safeName = System.Text.RegularExpressions.Regex.Replace(safeName, @"[""*:<>?/\\|\t]", "_");

            safeName += System.IO.Path.GetExtension(filename).Trim();

            return safeName;
        }

        private static string Translate(string term, Dictionary<string, string> translations)
        {
            if (translations.ContainsKey(term))
            {
                return translations[term];
            }
            return null;
        }

        private static Dictionary<string, string> Translation(string translateVar)
        {
            Dictionary<string, string> terms = new Dictionary<string, string>();
            string translateStr = System.Environment.GetEnvironmentVariable(translateVar);
            if (!string.IsNullOrEmpty(translateStr))
            {
                foreach (string item in translateStr.Split(";;", StringSplitOptions.RemoveEmptyEntries))
                {
                    string[] kvp = item.Split("::", StringSplitOptions.None);
                    if (!terms.ContainsKey(kvp[0]))
                    {
                        terms.Add(kvp[0], kvp[1]);
                    }
                }
            }
            return terms;
        }
    }
}