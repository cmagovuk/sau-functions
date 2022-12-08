using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace CMA.SAU.AzureFunctions
{
    internal class User
    {
        internal static async Task InviteToRoleAsync(ILogger log, Response response, string userEmail, string role)
        {
            log.LogInformation($"User.InviteToRole {userEmail} role: {role}");
            string groupId = Utilities.TranslateOne(role, "ROLE_MAPPINGS");

            if (!String.IsNullOrWhiteSpace(groupId) && !String.IsNullOrWhiteSpace(userEmail))
            {
                GraphServiceClient graphClient = Utilities.GetGraphClientWithCert();

                string userId = await SendInvitation(log, graphClient, userEmail);

                bool isGroupMember = await IsMemberOfGroup(graphClient, userId, groupId);

                if (!isGroupMember)
                {
                    await AddUserToGroup(graphClient, userId, groupId);
                }

                response.data = userId;
            }
            else
            {
                throw (new ArgumentException("Unknown role"));
            }
        }

        internal static async Task HasCaseAccessAsync(ILogger log, Response response, string userId, string requestId)
        {
            log.LogInformation($"User.HasCaseAccess {userId} requestId: {requestId}");
            string submission_list = System.Environment.GetEnvironmentVariable("SUBMISSIONS_LIST");
            string caseGroupId = null;

            using (ClientContext ctx = Utilities.GetSAUCasesContext())
            {
                Microsoft.SharePoint.Client.List list = ctx.Web.Lists.GetByTitle(submission_list);


                CamlQuery query = CamlQuery.CreateAllItemsQuery();
                query.ViewXml = "<View>" +
                        "<ViewFields><FieldRef Name='SAUCaseGroupID'/></ViewFields>" +
                         "<Query>" +
                         "<Where>" +
                         $"<Eq><FieldRef Name='SAURequestUniqueID' /><Value Type='Text'>{requestId}</Value></Eq>" +
                         "</Where>" +
                         "</Query>" +
                         "</View>";

                ListItemCollection items = list.GetItems(query);
                ctx.Load(items);
                ctx.ExecuteQueryRetry();

                if (items.Count == 1)
                {
                    caseGroupId = items[0]["SAUCaseGroupID"] as string;
                }
            }
            if (!string.IsNullOrEmpty(caseGroupId))
            {
                GraphServiceClient gc = Utilities.GetGraphClientWithCert();
                response.success = true;
                response.data = await IsMemberOfGroup(gc, userId, caseGroupId);
            }
            else
            {
                response.success = true;
                response.data = false;
                log.LogInformation($"Unable to find Case Group Id for request id {requestId}");
            }
            return;
        }
        private static async Task AddUserToGroup(GraphServiceClient graphClient, string userId, string groupId)
        {
            var directoryObject = new DirectoryObject
            {
                Id = $"{{{userId}}}"
            };

            await graphClient.Groups[$"{{{groupId}}}"].Members.References
                .Request()
                .AddAsync(directoryObject);
            // Use user Id to set up group membership result.InvitedUser.Id
            // Add userId to Response, so caller can keep reference
        }

        private static async Task<bool> IsMemberOfGroup(GraphServiceClient graphClient, string userId, string groupId)
        {
            var queryOptions = new List<QueryOption>()
                {
                    new QueryOption("$count", "true")
                };

            var groupMembersResult = await graphClient.Groups[$"{{{groupId}}}"].Members
                .Request(queryOptions)
                .Header("ConsistencyLevel", "eventual")
                .Filter($"id eq '{userId}'")
                .GetAsync();

            return groupMembersResult.Count != 0;
        }

        private static async Task<string> SendInvitation(ILogger log, GraphServiceClient graphClient, string userEmail)
        {
            try
            {
                string redirectUrl = Environment.GetEnvironmentVariable("PAP_URL");

                Invitation invitation = new Invitation
                {
                    InvitedUserEmailAddress = userEmail,
                    InviteRedirectUrl = redirectUrl,
                    SendInvitationMessage = true
                };
                var result = await graphClient.Invitations.Request().AddAsync(invitation);

                return result.InvitedUser.Id;
            }

            catch (Microsoft.Graph.ServiceException ex)
            {
                /*
                 * Invitation will fail with ServiceException if the email address is a verified domain of this directory
                 * Try to get the user Id by using the email address as the UPN of the user in the directory
                 */
                log.LogInformation(ex, "Invitation failed for {0}", userEmail);
                var result = await graphClient.Users[userEmail].Request().GetAsync();
                return result.Id;
            }
        }
    }
}