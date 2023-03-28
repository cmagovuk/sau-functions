using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace CMA.SAU.AzureFunctions
{
    public static class API
    {
        [FunctionName("API")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            Response response = new Response();

            try
            {
                // Get request body
                string reqBody = await new StreamReader(req.Body).ReadToEndAsync();
                Request request = new Request(reqBody);

                var payload = request.Payload;
                switch (request.Method)
                {
                    case "User.InviteToRole":
                        await User.InviteToRoleAsync(log, response, (string)payload.email, (string)payload.role);
                        break;

                    case "User.Remove":
                        await User.Remove(log, response, (string)payload.userId, (string)payload.role);
                        break;

                    case "User.HasCaseAccess":
                        await User.HasCaseAccessAsync(log, response, (string)payload.userId, (string)payload.requestId);
                        break;

                    case "User.PACreatedUser":
                        await User.PACreatedUser(log, response, (string)payload.userCreated, (string)payload.creatorName, (string)payload.creatorEmail, (string)payload.pa_name);
                        break;

                    case "RequestReport.Submit":
                        RequestReport.Submit(log, response, payload.reference, payload.uniqueId, payload.documents, payload.request);
                        break;

                    case "RequestReport.InformationResponse":
                        RequestReport.InformationResponse(log, response, payload.uniqueId, payload.documents);
                        break;

                    case "RequestReport.WithdrawRequest":
                        RequestReport.WithdrawRequest(log, response, payload.uniqueId, payload.documents);
                        break;

                    case "RequestReport.Mailbox":
                        RequestReport.Mailbox(log, response, payload.uniqueId);
                        break;

                    case "Test.Group":
                        await Test.Group(log, response, (string)payload.input);
                        break;

                    case "Test.Email":
                        Test.Email(log, response, (string)payload.email);
                        break;

                    default:
                        response.success = false;
                        response.error = "Unknown method";
                        break;
                }
            }
            catch (Exception ex)
            {
                response.success = false;
                response.error = ex.ToString();
            }
            log.LogInformation($"response.success: {response.success}");
            log.LogInformation($"response.error: {response.error}");
            return new OkObjectResult(response);

            //            string name = req.Query["name"];
            //
            //            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            //            dynamic data = JsonConvert.DeserializeObject(requestBody);
            //            name = name ?? data?.name;
            //
            //            string responseMessage = string.IsNullOrEmpty(name)
            //                ? "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."
            //                : $"Hello, {name}. This HTTP triggered function executed successfully.";
            //
            //            return new OkObjectResult(responseMessage);
        }
    }
}
