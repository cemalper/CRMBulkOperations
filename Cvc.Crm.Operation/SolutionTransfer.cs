using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Xml;
using Xrm;

namespace Cvv.Crm.Operation
{
    public static class SolutionTransfer
    {
        private static string GetPrefix(IOrganizationService service, string solutionName = null)
        {
            string prefix = string.Empty;
            var host = (service as OrganizationServiceProxy)?.ServiceManagement?.CurrentServiceEndpoint?.Address?.Uri?.Host ?? "Unknown";

            if (string.IsNullOrWhiteSpace(solutionName) == false)
            {
                prefix = $"Solution:{ solutionName}|";
            }
            return $"{prefix} Host:{host} =>";
        }

        private static List<string> GetSolutionNames(IOrganizationService service)
        {
            QueryExpression query = new QueryExpression(Solution.EntityLogicalName)
            {
                ColumnSet = new ColumnSet(Solution.Fields.UniqueName)
            };
            query.Criteria.AddCondition(Solution.Fields.IsManaged, ConditionOperator.Equal, false);
            query.Criteria.AddCondition(Solution.Fields.SolutionType, ConditionOperator.Equal, 0/*Hiçbiri*/);
            var solutionNames = service.RetrieveMultiple(query).Entities.Select(x => x.ToEntity<Solution>().UniqueName).ToList();

            return solutionNames.OrderBy(x => x).ToList();
        }

        public static byte[] ExportSolution(IOrganizationService service, string solutionName = "Default", string folder = "solution", bool isManaged = false, Action<string> callBack = null)
        {
            var host = (service as OrganizationServiceProxy)?.ServiceManagement?.CurrentServiceEndpoint?.Address?.Uri?.Host ?? "Unknown";
            callBack?.Invoke($"{GetPrefix(service, solutionName)} Export to byte is started.");
            Stopwatch sw = new Stopwatch();
            sw.Start();

            ExportSolutionRequest exportSolutionRequest = new ExportSolutionRequest
            {
                Managed = isManaged,
                SolutionName = solutionName
            };
            ExportSolutionResponse exportSolutionResponse = (ExportSolutionResponse)service.Execute(exportSolutionRequest);
            byte[] exportFile = exportSolutionResponse.ExportSolutionFile;
            sw.Stop();

            callBack?.Invoke($"{GetPrefix(service, solutionName)} Export to byte is finished. Folder: {folder}. Elapsed Time:{sw.Elapsed.ToString()}");

            return exportFile;
        }

        public static void ExportSolutionsToFile(IOrganizationService service, string folder = "solution", Action<string> callBack = null)
        {
            callBack?.Invoke($"{GetPrefix(service)} Export solutions is started.");
            var sw = new Stopwatch();
            sw.Start();
            var solutionNames = GetSolutionNames(service);
            solutionNames.ForEach(solutionName => ExportSolutionToFile(service, solutionName, folder, callBack));
            //Parallel.ForEach(solutionNames,new ParallelOptions { MaxDegreeOfParallelism = 4 }, solutionName => ExportSolutionToFile(service, solutionName, folder, callBack));
            sw.Stop();
            callBack?.Invoke($"{GetPrefix(service)} Export solutions is finished. Total Elapsedtime: {sw.Elapsed.ToString()}");
        }

        public static void ExportSolutionToFile(IOrganizationService service, string solutionName = "Default", string folder = "solution", Action<string> callBack = null)
        {
            callBack?.Invoke($"{GetPrefix(service, solutionName)} Export solution to file is started.");
            var solutionVersion = FindSolutionVersion(service, solutionName);
            var sw = new Stopwatch();
            sw.Start();
            var file = ExportSolution(service, solutionName, folder, false, callBack);
            var date = DateTime.Now.ToString("yyyyMMddhhmm");
            string fileName = $"{solutionName}_{solutionVersion}_{date}.zip";
            string path = $"{folder}/{fileName}";
            Directory.CreateDirectory(folder);
            File.WriteAllBytes(path, file);
            sw.Stop();
            callBack?.Invoke($"{GetPrefix(service, solutionName)} Export solution to file is finished. Path: {path}. Elapsedtime: {sw.Elapsed.ToString()}");
        }

        public static void ImportSolutionFromFile(List<IOrganizationService> services, string solutionName, string path, bool isManagedSolution = false, Action<string> callBack = null)
        {
            byte[] fileBytes = File.ReadAllBytes(path);

            ImportSolution(services, fileBytes, solutionName, isManagedSolution: isManagedSolution, traceCallback: callBack);
        }

        public static void ImportSolution(List<IOrganizationService> services, byte[] solutionFile, string solutionName = "Default", string xmlFolder = "ImportLogs", bool isManagedSolution = false, bool isPublish = false, Action<string> traceCallback = null)
        {
            Stopwatch sw = new Stopwatch();
            traceCallback?.Invoke($"Solution import tasks are started. Solution name: {solutionName}.");
            sw.Start();
            //services.ForEach(service =>
            Parallel.ForEach(services, service =>
            {
                string prefixInfo = GetPrefix(service, solutionName);

                var host = (service as OrganizationServiceProxy)?.ServiceManagement?.CurrentServiceEndpoint?.Address?.Uri?.Host ?? "Unknown";
                try
                {
                    var importJobId = Guid.NewGuid();
                    var importsolutionRequest = new ImportSolutionRequest()
                    {
                        ImportJobId = importJobId,
                        PublishWorkflows = true,
                        CustomizationFile = solutionFile,
                    };
                    if (isManagedSolution)
                    {
                        importsolutionRequest.OverwriteUnmanagedCustomizations = true;
                        importsolutionRequest.HoldingSolution = true;
                    }
                    ExecuteAsyncRequest asyncRequest = new ExecuteAsyncRequest
                    {
                        Request = importsolutionRequest,
                    };

                    traceCallback?.Invoke($"{prefixInfo} ImportSolution request is executing");
                    ExecuteAsyncResponse asyncResponse = (ExecuteAsyncResponse)service.Execute(asyncRequest);
                    AsyncOperation operation = null;
                    double? progress;
                    var result = Task.Run(() =>
                    {
                        do
                        {
                            operation = service.Retrieve(AsyncOperation.EntityLogicalName, asyncResponse.AsyncJobId, new ColumnSet(AsyncOperation.Fields.StateCode, AsyncOperation.Fields.StatusCode, AsyncOperation.Fields.FriendlyMessage)).ToEntity<AsyncOperation>();
                            try
                            {
                                progress = service.Retrieve(ImportJob.EntityLogicalName, importJobId, new ColumnSet(new string[] { ImportJob.Fields.Progress })).ToEntity<ImportJob>().Progress;
                                traceCallback?.Invoke($"{prefixInfo} Import Progress: {progress}");
                            }
                            catch (Exception)
                            {
                            }
                            if (operation.StateCode != AsyncOperationState.Completed)
                            {
                                Task.Delay(10000).Wait();
                            }
                        } while (operation.StateCode != AsyncOperationState.Completed);
                        progress = service.Retrieve(ImportJob.EntityLogicalName, importJobId, new ColumnSet(new string[] { ImportJob.Fields.Progress })).ToEntity<ImportJob>().Progress;
                        traceCallback?.Invoke($"{prefixInfo} Import Progress: {progress}");
                        return operation;
                    });
                    result.Wait();
                    operation = result.Result;
                    traceCallback?.Invoke($"{prefixInfo} ---------------------------------------------");
                    traceCallback?.Invoke($"{prefixInfo} Async operation results is {(operation.StatusCodeEnum ?? null)}, Friendly Message is {HttpUtility.HtmlDecode(operation.FriendlyMessage)}.");
                    traceCallback?.Invoke($"{prefixInfo} ---------------------------------------------");
                    traceCallback?.Invoke($"{prefixInfo} ImportSolution request is executed. ImportJob request is executing");
                    var sww = new Stopwatch();
                    sw.Start();
                    Task.Delay(10000).Wait();
                    var elapsedTime = sw.ElapsedMilliseconds;
                    traceCallback?.Invoke($"{prefixInfo} ElapsedTime: {elapsedTime}");
                    ImportJob job = service.Retrieve(ImportJob.EntityLogicalName, importJobId, new ColumnSet(new string[] { ImportJob.Fields.Data, ImportJob.Fields.SolutionName })).ToEntity<ImportJob>();
                    traceCallback?.Invoke($"{prefixInfo} ImportJob request is executed. ImportJob request is executed");

                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(job.Data);
                    try
                    {
                        Directory.CreateDirectory(xmlFolder);
                        var solutionVersion = FindSolutionVersion(service, solutionName);
                        var xmlFileName = $"{xmlFolder}/ImportResult_{solutionName}_{solutionVersion}_{host}_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xml";
                        traceCallback?.Invoke($"{prefixInfo} Importlog is saved in {xmlFileName}");
                        traceCallback?.Invoke($"{prefixInfo} Log result is {doc.SelectSingleNode("/importexportxml/solutionManifests/solutionManifest/result")?.OuterXml}");
                        doc.Save(xmlFileName);
                    }
                    catch (Exception)
                    {
                        traceCallback?.Invoke($"{prefixInfo} Import xml save is failed");
                    }
                    if (isManagedSolution)
                    {
                        traceCallback?.Invoke($"{prefixInfo} DeleteAndPromoteRequest is started");
                        var request = new DeleteAndPromoteRequest
                        {
                            UniqueName = solutionName
                        };
                        var response = (DeleteAndPromoteResponse)service.Execute(request);
                        traceCallback?.Invoke($"{prefixInfo} DeleteAndPromoteRequest is finished");
                    }
                    if (isPublish)
                    {
                        traceCallback?.Invoke($"{prefixInfo} Publishing...");
                        Publish(service);
                        traceCallback?.Invoke($"{prefixInfo} Publish is finished");
                    }
                }
                catch (Exception ex)
                {
                    traceCallback?.Invoke($"{solutionName} import is failed to {host}. { ex.ToString()}");
                }
            });

            sw.Stop();

            traceCallback?.Invoke($"Solution import tasks are finished {solutionName}. Elapsed Time:{sw.Elapsed.ToString()}");
        }

        public static bool ImportLanguage(OrganizationServiceProxy service, byte[] solutionFile, Action<string> traceCallback = null)
        {
            Stopwatch sw = new Stopwatch();
            traceCallback?.Invoke($"Language import tasks are started");
            sw.Start();
            //services.ForEach(service =>
            string prefixInfo = "Language";

            var host = (service as OrganizationServiceProxy)?.ServiceManagement?.CurrentServiceEndpoint?.Address?.Uri?.Host ?? "Unknown";
            try
            {
                var importJobId = Guid.NewGuid();

                var request = new ImportTranslationRequest()
                {
                    ImportJobId = importJobId,
                    TranslationFile = solutionFile
                };

                traceCallback?.Invoke($"{prefixInfo} ImportLanguageSolution request is executing");
                bool isCompletedSuccessfully = false;
                Task<ImportTranslationResponse> task = null;
                do
                {
                    task = Task.Factory.StartNew<ImportTranslationResponse>(() =>
                    {
                        return (ImportTranslationResponse)service.Execute(request);
                    });

                    isCompletedSuccessfully = task.Wait(TimeSpan.FromMilliseconds(1000 * 60 * 7));
                    if (!isCompletedSuccessfully)
                    {
                        traceCallback?.Invoke("Not Completed.Trying again");
                    }
                } while (!isCompletedSuccessfully);

                ImportTranslationResponse asyncResponse = task.Result;
                traceCallback?.Invoke($"{prefixInfo} ImportSolution request is executed. ImportJob request is executing");
                ImportJob job = service.Retrieve(ImportJob.EntityLogicalName, importJobId, new ColumnSet(new string[] { ImportJob.Fields.Data, ImportJob.Fields.SolutionName })).ToEntity<ImportJob>();
                traceCallback?.Invoke($"{prefixInfo} ImportJob request is executed. ImportJob request is executed");

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(job.Data);
                try
                {
                    var folder = "ImportLogs";
                    Directory.CreateDirectory(folder);
                    var xmlFileName = $"{folder}/ImportResult_lang_{host}_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xml";
                    traceCallback?.Invoke($"{prefixInfo} Importlog is saved in {xmlFileName}");
                    doc.Save(xmlFileName);
                }
                catch (Exception)
                {
                    traceCallback?.Invoke($"{prefixInfo} Import xml save is failed");

                    return false;
                }
            }
            catch (Exception ex)
            {
                traceCallback?.Invoke($"Lang import is failed to {host}. { ex.ToString()}");

                return false;
            }

            sw.Stop();

            traceCallback?.Invoke($"Solution import tasks are finished Lang. Elapsed Time:{sw.Elapsed.ToString()}");

            return true;
        }

        public static void TransferSolution(IOrganizationService fromService, List<IOrganizationService> toService, string solutionName = "Default", bool isManaged = false, bool isPublish = false, Action<string> callBack = null)
        {
            var file = ExportSolution(fromService, solutionName, isManaged: isManaged, callBack: callBack);
            ImportSolution(toService, solutionFile: file, solutionName: solutionName, isManagedSolution: isManaged, isPublish: isPublish, traceCallback: callBack);
        }

        public static void TransferSolutions(IOrganizationService fromService, List<IOrganizationService> toService, string[] solutionNames, bool isManaged = false, bool isPublish = false, Action<string> callBack = null)
        {
            var sw = new Stopwatch();
            sw.Start();
            callBack?.Invoke("Solution names are retrieving");
            callBack?.Invoke("Solution names are retrieved");
            Parallel.ForEach(solutionNames, solutionName =>
            ///foreach (var solutionName in solutionNames)
            {
                var file = ExportSolution(fromService, solutionName, callBack: callBack);
                ImportSolution(toService, solutionFile: file, solutionName: solutionName, isManagedSolution: isManaged, isPublish: isPublish, traceCallback: callBack);
            });
            sw.Stop();
            callBack?.Invoke($"Total elapsed time: {sw.ElapsedMilliseconds}");
        }

        public static void TransferSolutions(IOrganizationService fromService, List<IOrganizationService> toService, bool isManagadSolution = false, bool isPublish = false, Action<string> callBack = null)
        {
            var sw = new Stopwatch();
            sw.Start();
            callBack?.Invoke("Solution names are retrieving");
            var solutionNames = GetSolutionNames(fromService);
            callBack?.Invoke("Solution names are retrieved");
            foreach (var solutionName in solutionNames)
            {
                var file = ExportSolution(fromService, solutionName, callBack: callBack);
                ImportSolution(toService, solutionFile: file, solutionName: solutionName, isManagedSolution: isManagadSolution, isPublish: isPublish, traceCallback: callBack);
            }
            sw.Stop();
            callBack?.Invoke($"Total elapsed time: {sw.ElapsedMilliseconds}");
        }

        public static PublishAllXmlResponse Publish(IOrganizationService service)
        {
            PublishAllXmlRequest publishRequest = new PublishAllXmlRequest();
            return (PublishAllXmlResponse)service.Execute(publishRequest);
        }

        private static string FindSolutionVersion(IOrganizationService service, string solutionName)
        {
            var query = new QueryExpression
            {
                EntityName = Solution.EntityLogicalName,
                ColumnSet = new ColumnSet(Solution.Fields.Version),
                NoLock = true,
                TopCount = 1,
            };
            query.Criteria.AddCondition(Solution.Fields.UniqueName, ConditionOperator.Equal, solutionName);

            return service.RetrieveMultiple(query).Entities.ElementAt(0)?.GetAttributeValue<string>(Solution.Fields.Version);
        }
    }
}