using Cvc.Crm.Operation;
using Cvv.Crm.Operation;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Cvc.Crm.ConsoleApp
{
    class Program
    {
        static void Main()
        {
            ServicePointManager.DefaultConnectionLimit = 5000;
            ServicePointManager.MaxServicePointIdleTime = 5000;

            var connectionString = ConfigurationManager.ConnectionStrings["CrmConnection"].ConnectionString;
            CrmServiceClient crmSvc = new CrmServiceClient(connectionString);
            var response = crmSvc.OrganizationServiceProxy.Execute(new WhoAmIRequest());

            var requests = Enumerable.Range(1, 10000).Select(x =>
            {
                var entity = new Entity("contact");
                entity["firstname"] = $"First Name {x}";
                entity["lastname"] = $"Lastname";
                return new CreateRequest { Target = entity };
            }).ToList();

            MultipleExecutor<CreateRequest, CreateResponse>(() => crmSvc.OrganizationServiceProxy, requests);

            Console.ReadKey();
        }

        static void MultipleExecutor<TOrganizatioRequest, TOrganizationResponse>(Func<OrganizationServiceProxy> serviceCreator, List<TOrganizatioRequest> requests, int chunkSize = 80, int threadCount = 32) where TOrganizatioRequest : OrganizationRequest where TOrganizationResponse : OrganizationResponse
        {
            var bulkRequest = new CrmBulkRequestExecutor<TOrganizatioRequest, TOrganizationResponse>(requests, serviceCreator, new RequestOptions { ChuckSize = chunkSize, ParalelSize = threadCount });
            bulkRequest.ChuckCompleted += (obj, arg) =>
            {
                Console.WriteLine($"Remaing Request: {arg.RemainingRequestCount} Running Request: {arg.RunningRequestCount}");
            };
            bulkRequest.ErrorOccursed += (obj, arg) =>
            {
                Console.BackgroundColor = ConsoleColor.Red;
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine($"Error: {arg.Exception.Message}");
                Console.ResetColor();
            };
            bulkRequest.RequestErrorOccursed += (obj, arg) =>
            {
                Console.BackgroundColor = ConsoleColor.Red;
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine($"Request Error: {arg?.RequestContainer?.Exceptions?.Last()?.Message}");
                Console.ResetColor();
            };
            bulkRequest.ThreadFinished += (obj, arg) => { Console.WriteLine($"Remaining Thread Count: {arg.RemainingThreadCount}"); };
            bulkRequest.StatisticInfo += (obj, arg) =>
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Statistic: Avg RPS: {arg.AverageRequestCount}, FinishedRecord: {arg.FinishedRequestCount}, Total Time: {arg.RunningTime}, Percentage: {arg.ProgressPercentage}");
                Console.ResetColor();
            };
            bulkRequest.Completed += (obj, arg) =>
            {
                Console.WriteLine($"Completed Req:{arg.RequestsContainers.Count} : FailedReq: {arg.RequestsContainers.Count(x => x.IsFailed)}, Time {arg.ElapsedTime.ToString()} Avg: {arg.RequestsContainers.Count / arg.ElapsedTime.TotalSeconds}");
            };
            bulkRequest.Execute();
        }
    }
}
