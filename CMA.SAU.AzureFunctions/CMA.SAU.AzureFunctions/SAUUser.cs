using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Online.SharePoint.TenantAdministration;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Net.Mail;
using System.Threading.Tasks;

namespace CMA.SAU.AzureFunctions
{
    internal class SAUUser
    {
        internal static async Task AddUser(ILogger log, Response response, string userGroup, string email)
        {
            log.LogInformation($"SAUUser.AddUser userGroup: {userGroup} email: {email}");

            string groupId = GetGroupId(userGroup);
            GraphServiceClient graphClient = Utilities.GetGraphClientWithCert();

            string userId = await GetIdFromEmail(graphClient, email);
            bool isGroupMember = await IsMemberOfGroup(graphClient, userId, groupId);
            if (!isGroupMember)
            {
                await AddUserToGroup(graphClient, userId, groupId);
            }

            response.success = true;
            return;
        }

        internal static async Task GetUsers(ILogger log, Response response, string userGroup)
        {
            log.LogInformation($"SAUUser.GetUsers userGroup: {userGroup}");

            GraphServiceClient graphClient = Utilities.GetGraphClientWithCert();
            List<UserInfo> userInfo = await GetUsersForGroup(graphClient, GetGroupId(userGroup));

            response.success = true;
            response.data = userInfo;
        }

        internal static async Task GetAllUsers(ILogger log, Response response)
        {
            log.LogInformation($"SAUUser.GetAllUsers");
            AllUserInfo users = new();
            GraphServiceClient graphClient = Utilities.GetGraphClientWithCert();

            users.Admin = await GetUsersForGroup(graphClient, GetGroupId("admin"));
            users.Lead = await GetUsersForGroup(graphClient, GetGroupId("lead"));
            users.Team = await GetUsersForGroup(graphClient, GetGroupId("team"));

            response.success = true;
            response.data = users;
        }

        internal static async Task RemoveUser(ILogger log, Response response, string userGroup, string userId)
        {
            log.LogInformation($"SAUUser.AddUser userGroup: {userGroup} userId: {userId}");

            string groupId = GetGroupId(userGroup);
            GraphServiceClient graphClient = Utilities.GetGraphClientWithCert();

            //string userId = await GetIdFromEmail(graphClient, email);

            bool isGroupMember = await IsMemberOfGroup(graphClient, userId, groupId);
            if (isGroupMember)
            {
                await RemoveUserFromGroup(graphClient, userId, groupId);
            }
            response.success = true;
        }

        private static string GetGroupId(string userGroup)
        {
            switch (userGroup)
            {
                case "admin":
                    return Environment.GetEnvironmentVariable("SAU_ADMIN_GROUP_ID");
                case "lead":
                    return Environment.GetEnvironmentVariable("SAU_LEADER_GROUP_ID");
                case "team":
                    return Environment.GetEnvironmentVariable("SAU_TEAM_GROUP_ID");
            }
            return null;
        }

        private static async Task<List<UserInfo>> GetUsersForGroup(GraphServiceClient graphClient, string groupId)
        {
            List<UserInfo> userInfo = new();
            var groupMembersResult = await graphClient.Groups[$"{{{groupId}}}"].Members
                .Request()
                .GetAsync();

            foreach (DirectoryObject item in groupMembersResult.CurrentPage)
            {
                if (item is Microsoft.Graph.User user)
                {
                    userInfo.Add(new UserInfo(user));
                }
            }
            return userInfo;
        }

        private static async Task<string> GetIdFromEmail(GraphServiceClient graphClient, string userEmail)
        {
            // try getting user, using email as UPN
            try
            {
                var result = await graphClient.Users[userEmail].Request().GetAsync();
                return result.Id;
            }
            catch (ServiceException)
            {
                // most likely UPN is not email
                // try search users for email
                var result = await graphClient.Users.Request().Filter($"mail eq '{userEmail}'").GetAsync();
                if (result.Count == 1)
                {
                    return result[0].Id;
                }
            }
            return null;
        }

        private static async Task RemoveUserFromGroup(GraphServiceClient graphClient, string userId, string groupId)
        {
            await graphClient.Groups[$"{{{groupId}}}"].Members[$"{{{userId}}}"].Reference
                .Request()
                .DeleteAsync();
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
    }
}