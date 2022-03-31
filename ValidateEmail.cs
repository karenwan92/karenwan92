using System;
using System.IO;
using System.ServiceModel.Description;
using System.Threading.Tasks;
using System.Web.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using CAO.CRM.Common.Validation.DTO;
using CAO.CRM.Plugins.Common.Validation.Email;
using Microsoft.PowerPlatform.Dataverse.Client;

namespace FuncValidateEmailAddress
{
    public static class ValidateEmailAddress
    {
        [FunctionName("ValidateEmailAddress")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);

            // If input data is null, show block page
            if (data == null)
            {
                return (ActionResult)  new OkObjectResult(new ResponseContent("ShowBlockPage", "There was a problem with your request."));
            }

            // Print out the request body
            log.LogInformation("Request body: " + requestBody);
          
            // If email claim not found, show block page. Email is required and sent by default.
            if (data.email == null || data.email.ToString() == "" || data.email.ToString().Contains("@") == false)
            {
                return (ActionResult)new OkObjectResult(new ResponseContent("ShowBlockPage", "Email name is mandatory."));
            }

            string cao_url = GetEnvironmentVariable("CAO_OrganizationURL");
            string clientid = GetEnvironmentVariable("CRM_ClientId");
            string clientSecretid = GetEnvironmentVariable("CRM_ClientSecretId");

            ServiceClient service = GetServiceClient(clientid, clientSecretid, cao_url);


            bool valid = false;
            
            valid = Validate(data.email.ToString(), service, log);
            if (!valid)
                return new BadRequestObjectResult(new ResponseContent("ValidationError", "invalid email address"));

            return new OkObjectResult(new ResponseContent());
        }

        private static string GetEnvironmentVariable(string name) => Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
        private static bool Validate(string emailAddress, ServiceClient service, ILogger log)
        {

            EmailAdressToBeValidated validatedEmail = new EmailAdressToBeValidated { EmailAddress = emailAddress };
            EmailAddressValidator emailAddressValidator = new EmailAddressValidator(validatedEmail, service);
            ValidateOutput isValidAddress = emailAddressValidator.IsValidAsync().Result;

            if (!isValidAddress.IsSuccessfull)
            {
                log.LogInformation("The email is not valid or in a valid format.");
                return false;
                //throw new Exception("The email is not valid or in a valid format."); //Also caught in the front-end to render text using a Content Snippet
            }

            return true;
        }

        private static ServiceClient GetServiceClient(string clientId, string clientSecret, string environment)
        {
            var connectionString =
                @$"Url={environment};AuthType=ClientSecret;ClientId={clientId};ClientSecret={clientSecret};RequireNewInstance=true";

            ServiceClient serviceClient = new ServiceClient(connectionString);
            return serviceClient;
        }
    }
}
