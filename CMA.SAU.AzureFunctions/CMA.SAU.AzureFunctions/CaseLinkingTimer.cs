using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Azure.Storage.Blobs;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;

namespace CMA.SAU.AzureFunctions
{
    public class CaseLinkingTimer
    {
        [FunctionName("CaseLinkingTimer")]
        public void Run([TimerTrigger("%CASEWORK_LINK_CRON%")] TimerInfo myTimer, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
            ProcessCaseLinking(log);
        }

        private static void ProcessCaseLinking(ILogger log)
        {
            ListItemCollection items = GetCasesToLink();

            log.LogInformation($"Processing {items.Count} records");
            if (items.Count > 0)
            {
                string hubRequestLib = System.Environment.GetEnvironmentVariable("CASEWORK_REQUESTS_LIST");
                using ClientContext cc = Utilities.GetCaseworkHubContext();
                List requests = cc.Web.Lists.GetByTitle(hubRequestLib);
                foreach (ListItem sub in items)
                {
                    double? caseRequestId = sub[Constants.CASE_REQUEST_ID] as double?;

                    if (caseRequestId != null && caseRequestId != 0)
                    {
                        log.LogInformation($"Extracted Case Request Id: {caseRequestId}");
                        ListItem request = GetCaseRequest(requests, (double)caseRequestId, log);
                        if (request != null)
                        {
                            string status = request[Constants.STATUS] as string;
                            log.LogInformation($"Found status: {status}");

                            if (status != null)
                            {
                                if (status.ToLower().Contains("success"))
                                {
                                    log.LogInformation($"Site creation complete: {status}");

                                    if (request[Constants.SITE_URL] is string caseUrl)
                                    {
                                        log.LogInformation($"Found casework site {caseUrl}");
                                        ProcessCase(sub, caseUrl, log);
                                        UpdateSubmission(sub, request);
                                    }
                                    else
                                    {
                                        log.LogError("Failed to find Casework item");
                                    }
                                }
                                else
                                {
                                    log.LogError($"Failed to create case site. Status not Success: {status}");
                                }
                            }
                            else
                            {
                                log.LogInformation("Status returned null");
                            }
                        }
                    }
                    else
                    {
                        log.LogInformation($"Case Request Id is 0");
                    }
                }
            }
        }

        private static void ProcessCase(ListItem sub, string caseUrl, ILogger log)
        {
            using ClientContext ctx = Utilities.GetContext(caseUrl);
            CopyDocumentsToCaseSite(sub, ctx, log);
            SetFolderPermission(ctx, log);
            AddLinkToCaseSite(sub, ctx, log);
        }

        private static void SetFolderPermission(ClientContext ctx, ILogger log)
        {
            log.LogInformation($"Setting folder permissions");
            try
            {
                Folder folder = ctx.Web.EnsureFolderPath($"Shared Documents/PA Submission");
                ListItem folderItem = folder.ListItemAllFields;
                folderItem.ResetRoleInheritance();
                // folderItem.SystemUpdate();
                ctx.ExecuteQueryRetry();

                folderItem.BreakRoleInheritance(true, false);
                folderItem.EnsureProperty(i => i.RoleAssignments);
                foreach (RoleAssignment roleAssignment in folderItem.RoleAssignments)
                {
                    roleAssignment.EnsureProperty(a => a.Member);
                    roleAssignment.Member.EnsureProperty(a => a.LoginName);
                    folderItem.AddPermissionLevelToGroup(roleAssignment.Member.LoginName, RoleType.Reader, true);
                }

            }
            catch (Exception ex)
            {
                log.LogError(ex, "Failed during folder permissions");
            }
        }

        private static void CopyDocumentsToCaseSite(ListItem sub, ClientContext ctx, ILogger log)
        {
            string docsJSON = sub[Constants.DOCUMENT_JSON] as string;
            dynamic data = JsonConvert.DeserializeObject(docsJSON);
            log.LogInformation("Adding assessment documents");
            Utilities.CopyDocumentsToCaseSiteSubFolder(ctx, log, data.assessment_docs, "Assessment of compliance");
            log.LogInformation("Adding call in documents");
            Utilities.CopyDocumentsToCaseSiteSubFolder(ctx, log, data.call_in_docs, "Call in");
            log.LogInformation("Adding post award referral documents");
            Utilities.CopyDocumentsToCaseSiteSubFolder(ctx, log, data.par_docs, "Post Award Referral");
            log.LogInformation("Adding eligibility description documents");
            Utilities.CopyDocumentsToCaseSiteSubFolder(ctx, log, data.description_docs, "Eligibility description");
            log.LogInformation("Adding submission text documents");
            Utilities.CopyDocumentsToCaseSiteSubFolder(ctx, log, data.submission_docs, "Summary of submission");
        }

        private static void AddLinkToCaseSite(ListItem sub, ClientContext ctx, ILogger log)
        {
            log.LogInformation("Adding link to PAP");
            string pap_url = Environment.GetEnvironmentVariable("PAP_URL");
            string requestUniqueId = sub[Constants.UNIQUE_ID] as string;
            string title = $"SAU{sub["Title"]}";
            NavigationNodeCollection ql = ctx.Web.Navigation.QuickLaunch;
            ctx.Load(ql);
            ctx.ExecuteQueryRetry();

            NavigationNodeCreationInformation newNode = new NavigationNodeCreationInformation
            {
                Title = title,
                Url = $"{pap_url}/sau_requests/{requestUniqueId}"
            };

            NavigationNode prev = ql.Where(n => n.Title == "Documents").FirstOrDefault();
            if (prev != null)
            {
                newNode.PreviousNode = prev;
            }
            else newNode.AsLastNode = true;

            ql.Add(newNode);
            ctx.ExecuteQueryRetry();
        }

        private static void UpdateSubmission(ListItem sub, ListItem caseRequest)
        {
            sub[Constants.SAU_CASE_GROUP_ID] = caseRequest[Constants.GROUP_ID];
            sub[Constants.SAU_CASE_SITE_URL] = caseRequest[Constants.SITE_URL];
            sub[Constants.SAU_EXTERNAL_MAILBOX_ID] = caseRequest[Constants.EXTERNAL_MAILBOX_ID];
            sub[Constants.SAU_INTERNAL_MAILBOX_ID] = caseRequest[Constants.INTERNAL_MAILBOX_ID];
            sub.Update();
            sub.Context.ExecuteQueryRetry();
        }

        private static ListItemCollection GetCasesToLink()
        {
            string submission_list = System.Environment.GetEnvironmentVariable("SUBMISSIONS_LIST");

            using (ClientContext ctx = Utilities.GetSAUCasesContext())
            {
                List list = ctx.Web.Lists.GetByTitle(submission_list);


                CamlQuery query = CamlQuery.CreateAllItemsQuery();
                query.ViewXml = "<View>" +
                         "<Query>" +
                         "<Where>" +
                         $"<IsNull><FieldRef Name='SAUCaseSiteUrl' /></IsNull>" +
                         "</Where>" +
                         "</Query>" +
                         "</View>";

                ListItemCollection items = list.GetItems(query);
                ctx.Load(items);
                ctx.ExecuteQueryRetry();

                return items;
            }
        }

        private static ListItem GetCaseRequest(List requests, double caseRequestId, ILogger log)
        {
            try
            {
                ListItem request = requests.GetItemById(caseRequestId.ToString());
                requests.Context.Load(request);
                requests.Context.ExecuteQueryRetry();
                return request;
            }
            catch (Exception ex)
            {
                log.LogError($"Failed to load case request {caseRequestId}.  Error: {ex.Message}", ex);
                return null;
            }
        }
    }
}
