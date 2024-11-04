using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;

namespace CMA.SAU.AzureFunctions
{
    public class TeamAssignedTimer
    {
        [FunctionName("TeamAssignedTimer")]
        public void Run([TimerTrigger("%TEAM_ASSIGNED_CRON%")] TimerInfo myTimer, ILogger log)
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-gb");
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
            log.LogInformation(System.TimeZoneInfo.Local.DisplayName);
            try
            {
                ProcessTeamAssigned(log);
            }
            catch (Exception ex)
            {
                log.LogError(ex, "Error during team assigned timer job");
            }
        }

        private void ProcessTeamAssigned(ILogger log)
        {
            List<TeamAssignedInfo> requests = new List<TeamAssignedInfo>();

            ListItemCollection items = GetCasesToCheck(log);

            log.LogInformation($"Processing {items.Count} records");
            if (items.Count > 0)
            {
                foreach (Microsoft.SharePoint.Client.ListItem item in items)
                {
                    List<String> initialOwners = GetUserEmails(item[Constants.OWNERS] as FieldUserValue[]);
                    List<String> initialMembers = GetUserEmails(item[Constants.MEMBERS] as FieldUserValue[]);
                    TeamAssignedInfo info = new TeamAssignedInfo()
                    {
                        Changed = false,
                        RefNo = item["Title"] as string,
                        Url = item[Constants.SITE_URL] as string,
                        Created = (DateTime) item["Created"]
                    };
                    requests.Add(info);

                    string caseGroupId = item[Constants.GROUP_ID] as string;
                    if (!string.IsNullOrEmpty(caseGroupId))
                    {
                        Microsoft.Graph.GraphServiceClient gc = Utilities.GetGraphClientWithCert();
                        var groupMembers = Utilities.GetGroupMembers(gc, caseGroupId);
                        var groupOwners = Utilities.GetGroupOwners(gc, caseGroupId);

                        info.Changed = (DifferentUsers(initialOwners, groupOwners) || DifferentMembers(initialOwners, initialMembers, groupMembers));
                    }
                    else
                    {
                        info.Message = "Unable to determine site group, this could be due to an error during site creation or the request has just been submitted.";
                    }
                }
                SendEmail(requests, log);
            }
        }

        private void SendEmail(List<TeamAssignedInfo> requests, ILogger log)
        {
            List<TeamAssignedInfo> unchanged = requests.FindAll(r => !r.Changed);
            // Get email address to send to
            string recipients = System.Environment.GetEnvironmentVariable("TEAM_ASSIGNED_RECIPIENTS");
            string emailSubject = System.Environment.GetEnvironmentVariable("TEAM_ASSIGNED_SUBJECT");
            if (unchanged.Count > 0 && !string.IsNullOrWhiteSpace(recipients))
            {
                using ClientContext ctx = Utilities.GetSAUCasesContext();
                // Build email body
                List<string> tableData = new List<string>();
                foreach (TeamAssignedInfo tai in unchanged)
                {
                    string refNo = tai.Url != null ? $"<a href='{tai.Url}'>{tai.RefNo}</a>" : tai.RefNo;
                    tableData.Add($"<tr><td>{refNo}</td>"
                        + $"<td>{tai.Created:g}</td>"
                        + $"<td>{tai.Message ?? "Team not assigned"}</td>"
                        + "</tr>");
                }

                string emailBody = $"<p>Hi</p>"
                    + "<style>td {padding: 0px 10px 0px 10px}</style>" +
                    "<p>Listed below are recent SAU requests that have not been assigned teams 1 day after submission</p>"
                    + "<table>"
                    + "<tr>"
                    + "<th>Case ID</th>"
                    + "<th>Submitted</th>"
                    + "<th>Additional information</th>"
                    + "</tr>"
                    + string.Join("", tableData)
                    + "</table>";

                // Send email
                log.LogInformation($"Sending email to: {recipients}");
                Utilities.SendEmail(ctx, recipients, emailBody, emailSubject);
            }
        }

        private bool DifferentMembers(List<string> initialOwners, List<string> initialMembers, List<string> groupMembers)
        {
            // Merge iniital owners and members, removing any duplicates
            List<string> mergedUsers = initialOwners.Union(initialMembers).ToList();

            return DifferentUsers(mergedUsers, groupMembers);
        }

        private List<string> GetUserEmails(FieldUserValue[] userValues)
        {
            List<string> ids = new List<string>();
            foreach (var uv in userValues)
            {
                if (!ids.Contains(uv.Email.ToLower())) ids.Add(uv.Email.ToLower());
            }
            return ids;
        }

        private bool DifferentUsers(List<string> initialUsers, List<string> groupUsers)
        {
            if (initialUsers.Count != groupUsers.Count) return true;

            List<string> test = new List<string>(initialUsers.Intersect(groupUsers));

            return test.Count != initialUsers.Count;
        }

        private ListItemCollection GetCasesToCheckOld(ClientContext ctx, ILogger log)
        {
            string submission_list = System.Environment.GetEnvironmentVariable("SUBMISSIONS_LIST");

            Microsoft.SharePoint.Client.List list = ctx.Web.Lists.GetByTitle(submission_list);

            CamlQuery query = CamlQuery.CreateAllItemsQuery();
            query.ViewXml = "<View>" +
                     "<Query>" +
                     "<Where>" +
                     $"<Geq><FieldRef Name='Created' /><Value Type='DateTime'><Today OffsetDays='{OffsetDays()}' /></Value></Geq>" +
                     "</Where>" +
                     "</Query>" +
                     "</View>";

            ListItemCollection items = list.GetItems(query);
            ctx.Load(items);
            ctx.ExecuteQueryRetry();
            return items;
        }

        private ListItemCollection GetCasesToCheck(ILogger log)
        {
            string submission_list = System.Environment.GetEnvironmentVariable("CASEWORK_REQUESTS_LIST");
            string projectId = System.Environment.GetEnvironmentVariable("SAU_PROJECT_TYPE_ID");

            using ClientContext ctx = Utilities.GetCaseworkHubContext();
            Microsoft.SharePoint.Client.List list = ctx.Web.Lists.GetByTitle(submission_list);

            CamlQuery query = CamlQuery.CreateAllItemsQuery();
            query.ViewXml = "<View>" +
                     "<Query>" +
                     "<Where>" +
                     "<And>" +
                     $"<Eq><FieldRef Name='ProjectType' LookupId='true' /><Value Type='Lookup'>{projectId}</Value></Eq>" + 
                     $"<Geq><FieldRef Name='Created' /><Value Type='DateTime'><Today OffsetDays='{OffsetDays()}' /></Value></Geq>" +
                     "</And>" +
                     "</Where>" +
                     "</Query>" +
                     "</View>";

            ListItemCollection items = list.GetItems(query);
            ctx.Load(items);
            ctx.ExecuteQueryRetry();
            return items;
        }

        private string OffsetDays()
        {
            string offsetDaysEnv = System.Environment.GetEnvironmentVariable("TEAM_ASSIGNED_DAYS");
            if (!int.TryParse(offsetDaysEnv, out int offsetDays)) offsetDays = -10;

            return offsetDays.ToString();
        }
    }
}
