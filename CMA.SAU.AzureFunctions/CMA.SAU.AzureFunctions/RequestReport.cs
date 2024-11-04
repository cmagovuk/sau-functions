using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using System;

namespace CMA.SAU.AzureFunctions
{
    internal class RequestReport
    {
        internal static void Submit(ILogger log, Response response, dynamic reference, dynamic uniqueId, dynamic documents, dynamic request, dynamic projectName)
        {
            string submission_list = System.Environment.GetEnvironmentVariable("SUBMISSIONS_LIST");
            using (ClientContext ctx = Utilities.GetSAUCasesContext())
            {
                List list = ctx.Web.Lists.GetByTitle(submission_list);
                ListItemCreationInformation lici = new ListItemCreationInformation();
                ListItem listItem = list.AddItem(lici);
                listItem[Constants.TITLE] = ((string)reference).Replace("SAU","");
                listItem[Constants.UNIQUE_ID] = (string)uniqueId;
                listItem[Constants.DOCUMENT_JSON] = documents.ToString();
                listItem[Constants.REQUEST_JSON] = request?.ToString();
                listItem[Constants.SAU_PROJ_NAME] = (string)projectName;
                listItem.Update();
                ctx.ExecuteQueryRetry();
            }
        }

        internal static void InformationResponse(ILogger log, Response response, dynamic uniqueId, dynamic documents)
        {
            /*
             * Use unique id to find case site - using submissions list
             * Add response document to site
             */
            string submission_list = System.Environment.GetEnvironmentVariable("SUBMISSIONS_LIST");
            using ClientContext ctx = Utilities.GetSAUCasesContext();
            List list = ctx.Web.Lists.GetByTitle(submission_list);

            CamlQuery query = CamlQuery.CreateAllItemsQuery();
            query.ViewXml = "<View><ViewFields>" +
                    $"<FieldRef Name='{Constants.TITLE}'/><FieldRef Name='{Constants.SAU_CASE_GROUP_ID}'/>" +
                    $"<FieldRef Name='{Constants.SAU_CASE_SITE_URL}'/><FieldRef Name='{Constants.SAU_PROJ_NAME}'/>" +
                        "</ViewFields><Query>" +
                            "<Where>" +
                                $"<Eq><FieldRef Name='SAURequestUniqueID' /><Value Type='Text'>{uniqueId}</Value></Eq>" +
                            "</Where>" +
                         "</Query>" +
                     "</View>";

            ListItemCollection items = list.GetItems(query);
            ctx.Load(items);
            ctx.ExecuteQueryRetry();

            if (items.Count == 1)
            {
                ProcessResponse(log, uniqueId, documents);
                SendRFIResponseEmail(log, items[0]);
            }
        }

        internal static void WithdrawRequest(ILogger log, Response response, dynamic uniqueId, dynamic documents)
        {
            log.LogInformation($"RequestReport.WithdrawRequest uniqueId: {uniqueId}");
            ProcessDocumentUpload(log, response, uniqueId, documents, "Withdraw request");
        }

        internal static void ProcessDocumentUpload(ILogger log, Response response, dynamic uniqueId, dynamic documents, string folder)
        {
            /*
             * Use unique id to find case site - using submissions list
             * Add response document to site
             */
            string submission_list = System.Environment.GetEnvironmentVariable("SUBMISSIONS_LIST");
            using ClientContext ctx = Utilities.GetSAUCasesContext();
            List list = ctx.Web.Lists.GetByTitle(submission_list);

            CamlQuery query = CamlQuery.CreateAllItemsQuery();
            query.ViewXml = "<View><ViewFields>" +
                    $"<FieldRef Name='{Constants.TITLE}'/><FieldRef Name='{Constants.SAU_CASE_GROUP_ID}'/>" +
                    $"<FieldRef Name='{Constants.SAU_CASE_SITE_URL}'/><FieldRef Name='{Constants.SAU_PROJ_NAME}'/>" +
                     "</ViewFields><Query>" +
                     "<Where>" +
                     $"<Eq><FieldRef Name='SAURequestUniqueID' /><Value Type='Text'>{uniqueId}</Value></Eq>" +
                     "</Where>" +
                     "</Query>" +
                     "</View>";

            ListItemCollection items = list.GetItems(query);
            ctx.Load(items);
            ctx.ExecuteQueryRetry();

            if (items.Count == 1)
            {
                if (items[0][Constants.SAU_CASE_SITE_URL] is string caseUrl)
                {
                    log.LogInformation($"Found casework site {caseUrl}");
                    UploadDocuments(caseUrl, documents, log, folder);
                    SendWithdrawRequestEmail(log, items[0]);
                }
                else
                {
                    log.LogError("Failed to find Casework item");
                }
            }
        }

        private static void UploadDocuments(string caseUrl, dynamic documents, ILogger log, string folder)
        {
            using ClientContext ctx = Utilities.GetContext(caseUrl);
            Utilities.CopyDocumentsToCaseSiteSubFolder(ctx, log, documents, folder);
        }

        internal static void Mailbox(ILogger log, Response response, dynamic uniqueId)
        {
            log.LogInformation($"RequestReport.Mailbox uniqueId: {uniqueId}");
            string submission_list = System.Environment.GetEnvironmentVariable("SUBMISSIONS_LIST");
            using (ClientContext ctx = Utilities.GetSAUCasesContext())
            {
                List list = ctx.Web.Lists.GetByTitle(submission_list);

                CamlQuery query = CamlQuery.CreateAllItemsQuery();
                query.ViewXml = "<View>" +
                        $"<ViewFields><FieldRef Name='{Constants.SAU_EXTERNAL_MAILBOX_ID}'/></ViewFields>" +
                         "<Query>" +
                         "<Where>" +
                         $"<Eq><FieldRef Name='SAURequestUniqueID' /><Value Type='Text'>{uniqueId}</Value></Eq>" +
                         "</Where>" +
                         "</Query>" +
                         "</View>";

                ListItemCollection items = list.GetItems(query);
                ctx.Load(items);
                ctx.ExecuteQueryRetry();

                if (items.Count == 1)
                {
                    if (items[0][Constants.SAU_EXTERNAL_MAILBOX_ID] is string mailboxId)
                    {
                        Microsoft.Graph.GraphServiceClient graphClient = Utilities.GetGraphClientWithCert();
                        var result = graphClient.Users[mailboxId].Request().GetAsync().Result;
                        response.data = result.UserPrincipalName;
                    }
                    else
                    {
                        log.LogError("Failed to find Casework item");
                    }
                }
            }
        }
        private static void SendRFIResponseEmail(ILogger log, ListItem submissionRequest)
        {
            string caseGroupId = (string)submissionRequest[Constants.SAU_CASE_GROUP_ID];
            string caseUrl = (string)submissionRequest[Constants.SAU_CASE_SITE_URL];
            if (!string.IsNullOrEmpty(caseGroupId) && !string.IsNullOrEmpty(caseUrl))
            {
                string projectName = (string)submissionRequest[Constants.SAU_PROJ_NAME];
                string emailSubject = System.Environment.GetEnvironmentVariable("RFI_RESPONSE_SUBJECT");
                emailSubject = emailSubject.Replace("{ID}", $"SAU{submissionRequest[Constants.TITLE]}");
                // Build email body
                string emailBody =
                     $"<p>There has been a response to a request for information on SAU{submissionRequest[Constants.TITLE]}.</p>"
                    + $"<p>Check the response in the <a href='{caseUrl}'>{(projectName != null ? projectName : "SAU Request")} {submissionRequest[Constants.TITLE]}</a> SharePoint case site.</p>";

                SendRequestSiteEmail(log, caseUrl, caseGroupId, emailSubject, emailBody);
            }
        }

        private static void SendWithdrawRequestEmail(ILogger log, ListItem submissionRequest)
        {
            string caseGroupId = (string)submissionRequest[Constants.SAU_CASE_GROUP_ID];
            string caseUrl = (string)submissionRequest[Constants.SAU_CASE_SITE_URL];
            if (!string.IsNullOrEmpty(caseGroupId) && !string.IsNullOrEmpty(caseUrl))
            {
                string projectName = (string)submissionRequest[Constants.SAU_PROJ_NAME];
                string emailSubject = System.Environment.GetEnvironmentVariable("WITHDRAW_REQUEST_SUBJECT");
                emailSubject = emailSubject.Replace("{ID}", $"SAU{submissionRequest[Constants.TITLE]}");

                // Build email body
                string emailBody =
                     $"<p>There has been a request to withdraw SAU{submissionRequest[Constants.TITLE]}.</p>"
                    + $"<p>Check the withdraw document in the <a href='{caseUrl}'>{(projectName != null ? projectName : "SAU Request")} {submissionRequest[Constants.TITLE]}</a> SharePoint case site.</p>";

                SendRequestSiteEmail(log, caseUrl, caseGroupId, emailSubject, emailBody);
            }
        }

        private static void SendRequestSiteEmail(ILogger log, string caseUrl, string caseGroupId, string emailSubject, string emailBody)
        {
            try
            {
                Microsoft.Graph.GraphServiceClient gc = Utilities.GetGraphClientWithCert();
                var groupOwners = Utilities.GetGroupOwners(gc, caseGroupId);

                if (groupOwners.Count > 0)
                {
                    using ClientContext ctx = Utilities.GetContext(caseUrl);
                    // Send email
                    log.LogInformation($"Sending email to: {string.Join("; ", groupOwners)}");
                    Utilities.SendEmail(ctx, groupOwners, emailBody, emailSubject);
                }

            }
            catch (Exception ex)
            {
                log.LogError(ex, "Error sending email");
            }
        }

        private static void ProcessResponse(ILogger log, dynamic uniqueId, dynamic documents)
        {
            string submission_list = System.Environment.GetEnvironmentVariable("RFI_RESPONSES_LIST");
            using ClientContext ctx = Utilities.GetSAUCasesContext();
            List list = ctx.Web.Lists.GetByTitle(submission_list);
            ListItemCreationInformation lici = new ListItemCreationInformation();
            ListItem listItem = list.AddItem(lici);
            listItem[Constants.TITLE] = $"RFI Response for: {uniqueId}";
            listItem[Constants.UNIQUE_ID] = (string)uniqueId;
            listItem[Constants.DOCUMENT_JSON] = documents.ToString();
            listItem["SAUCompleted"] = false;
            listItem.Update();
            ctx.ExecuteQueryRetry();
        }
    }
}
