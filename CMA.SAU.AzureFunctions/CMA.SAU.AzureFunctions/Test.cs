using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace CMA.SAU.AzureFunctions
{
    struct TestResult
    {
        public int groupCount;
        public string inputString;
    }
    class Test
    {
        internal static async Task Group(ILogger log, Response response, string input)
        {
            GraphServiceClient graphClient = Utilities.GetGraphClientWithCert();


            var result = await graphClient.Groups.Request().Select("displayName").GetAsync();

            foreach (Microsoft.Graph.Group group in result.CurrentPage)
            {
                log.LogInformation(group.DisplayName);
            }

            TestResult testResult = new TestResult
            {
                groupCount = result.Count,
                inputString = input
            };

            response.data = testResult;

            // Use user Id to set up group membership result.InvitedUser.Id
            // Add userId to Response, so caller can keep reference
        }

        internal static void Email(ILogger log, Response response, string email)
        {
            using ClientContext ctx = Utilities.GetSAUCasesContext();
            List<string> recipients = new List<string> { email };

            var ep = new EmailProperties
            {
                To = recipients,
                Subject = "Test email",
                Body = "<p><b>Test</b> email</p>"
            };

            Utility.SendEmail(ctx, ep);
            ctx.ExecuteQueryRetry();
        }
    }
}
