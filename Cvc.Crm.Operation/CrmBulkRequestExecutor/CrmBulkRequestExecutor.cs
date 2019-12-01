using Cvc.Crm.Operation.Events;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.ServiceModel;
using System.Threading;
using System.Threading.Tasks;

namespace Cvc.Crm.Operation
{
    public class CrmBulkRequestExecutor<TOrganizatioRequest, TOrganizationResponse> where TOrganizatioRequest : OrganizationRequest where TOrganizationResponse : OrganizationResponse
    {
        #region Private Variables
        private const int _executeMultipleRequestCount = 10;
        private readonly Func<OrganizationServiceProxy> _serviceCreator;
        private readonly ConcurrentBag<RequestContainer<TOrganizatioRequest, TOrganizationResponse>> _requests;
        private readonly RequestOptions _requestOptions;
        private readonly Thread[] _tasks;

        private readonly object _takeAndRemoveLockObject = new object();
        private readonly object _chuckCompletedEventKey = new object();
        private readonly object _completedEventKey = new object();
        private readonly object _threadFinishedEventKey = new object();
        private readonly object _requestErrorOccuredEventKey = new object();
        private readonly object _errorOccuredEventKey = new object();
        private readonly object _statisticEventKey = new object();
        private readonly EventHandlerList _eventHandlerList = new EventHandlerList();
        private readonly System.Timers.Timer _statisticTimer = new System.Timers.Timer();
        private readonly Stopwatch _executionSw = new Stopwatch();
        #endregion Private Variables

        public List<RequestContainer<TOrganizatioRequest, TOrganizationResponse>> FailedRequests { get; private set; }

        public CrmBulkRequestExecutor(List<TOrganizatioRequest> requests, Func<OrganizationServiceProxy> serviceCreator, RequestOptions requestOptions = null)
        {
            //ServicePointManager.DefaultConnectionLimit = 5000;
            //ServicePointManager.MaxServicePointIdleTime = 5000;
            _requestOptions = requestOptions ?? new RequestOptions();
            _requests = new ConcurrentBag<RequestContainer<TOrganizatioRequest, TOrganizationResponse>>(requests.Select(x => new RequestContainer<TOrganizatioRequest, TOrganizationResponse>(x, _requestOptions.RetryCount)));
            _serviceCreator = serviceCreator;
            _tasks = new Thread[_requestOptions.ParalelSize];
            FailedRequests = new List<RequestContainer<TOrganizatioRequest, TOrganizationResponse>>();
            _statisticTimer.Interval = requestOptions.StatisticInverval;
            _statisticTimer.Elapsed += (sender, even) => CalculateStatistic();
        }

        #region Events

        private void OnChuckCompleted(ExecutorEventArg<TOrganizatioRequest, TOrganizationResponse> eventArgs)
        {
            var eventDelegate = (EventHandler<ExecutorEventArg<TOrganizatioRequest, TOrganizationResponse>>)_eventHandlerList[_chuckCompletedEventKey];
            eventDelegate?.Invoke(this, eventArgs);
        }

        private void OnCompleted(CompletedEventArgs<TOrganizatioRequest, TOrganizationResponse> eventArgs)
        {
            var eventDelegate = (EventHandler<CompletedEventArgs<TOrganizatioRequest, TOrganizationResponse>>)_eventHandlerList[_completedEventKey];
            eventDelegate?.Invoke(this, eventArgs);
        }

        private void OnThreadFinished(ThreadCompletedEventArg eventArgs)
        {
            var eventDelegate = (EventHandler<ThreadCompletedEventArg>)_eventHandlerList[_threadFinishedEventKey];
            eventDelegate?.Invoke(this, eventArgs);
        }

        private void OnRequestErrorOccurs(RequestEventArg<TOrganizatioRequest, TOrganizationResponse> eventArgs)
        {
            var eventDelegate = (EventHandler<RequestEventArg<TOrganizatioRequest, TOrganizationResponse>>)_eventHandlerList[_requestErrorOccuredEventKey];
            eventDelegate?.Invoke(this, eventArgs);
        }

        private void OnErrorOccurs(ErrorEventArg eventArgs)
        {
            var eventDelegate = (EventHandler<ErrorEventArg>)_eventHandlerList[_errorOccuredEventKey];
            eventDelegate?.Invoke(this, eventArgs);
        }

        private void OnStatisticInfo(StatisticEventArg eventArgs)
        {
            var eventDelegate = (EventHandler<StatisticEventArg>)_eventHandlerList[_statisticEventKey];
            eventDelegate?.Invoke(this, eventArgs);
        }

        public event EventHandler<ExecutorEventArg<TOrganizatioRequest, TOrganizationResponse>> ChuckCompleted
        {
            add
            {
                _eventHandlerList.AddHandler(_chuckCompletedEventKey, value);
            }
            remove
            {
                _eventHandlerList.RemoveHandler(_chuckCompletedEventKey, value);
            }
        }

        public event EventHandler<CompletedEventArgs<TOrganizatioRequest, TOrganizationResponse>> Completed
        {
            add
            {
                _eventHandlerList.AddHandler(_completedEventKey, value);
            }
            remove
            {
                _eventHandlerList.RemoveHandler(_completedEventKey, value);
            }
        }

        public event EventHandler<ThreadCompletedEventArg> ThreadFinished
        {
            add
            {
                _eventHandlerList.AddHandler(_threadFinishedEventKey, value);
            }
            remove
            {
                _eventHandlerList.RemoveHandler(_threadFinishedEventKey, value);
            }
        }

        public event EventHandler<ErrorEventArg> ErrorOccursed
        {
            add
            {
                _eventHandlerList.AddHandler(_errorOccuredEventKey, value);
            }
            remove
            {
                _eventHandlerList.RemoveHandler(_errorOccuredEventKey, value);
            }
        }

        public event EventHandler<RequestEventArg<TOrganizatioRequest, TOrganizationResponse>> RequestErrorOccursed
        {
            add
            {
                _eventHandlerList.AddHandler(_requestErrorOccuredEventKey, value);
            }
            remove
            {
                _eventHandlerList.RemoveHandler(_requestErrorOccuredEventKey, value);
            }
        }

        public event EventHandler<StatisticEventArg> StatisticInfo
        {
            add
            {
                _eventHandlerList.AddHandler(_statisticEventKey, value);
            }
            remove
            {
                _eventHandlerList.RemoveHandler(_statisticEventKey, value);
            }
        }

        #endregion Events

        public Task Execute(CancellationToken cancellationToken = default(CancellationToken))
        {
            return Task.Run(() =>
            {
                InternalExecute(cancellationToken);
            }, cancellationToken);
        }

        private void InternalExecute(CancellationToken cancellationToken = default(CancellationToken))
        {
            _executionSw.Restart();
            _statisticTimer.Start();
            for (int i = 0; i < _requestOptions.ParalelSize; i++)
            {
                var threadIndex = i;
                _tasks[i] = new Thread(() => ExecuteRequests(_serviceCreator, cancellationToken))
                {
                    Name = threadIndex.ToString()
                };
            }
            for (int i = 0; i < _requestOptions.ParalelSize; i++)
            {
                _tasks[i].Start();
            }
            for (int i = 0; i < _requestOptions.ParalelSize; i++)
            {
                _tasks[i].Join();
            }
            _executionSw.Stop();
            _statisticTimer.Start();
            var averageRecordCount = _requests.Count(x => x.IsCompleted) / _executionSw.Elapsed.TotalSeconds;
            OnCompleted(new CompletedEventArgs<TOrganizatioRequest, TOrganizationResponse>(_requests.ToList(), _executionSw.Elapsed, averageRecordCount));
        }

        private List<RequestContainer<TOrganizatioRequest, TOrganizationResponse>> TakeRequests(IEnumerable<RequestContainer<TOrganizatioRequest, TOrganizationResponse>> _requests, int? chunkSize = null)
        {
            List<RequestContainer<TOrganizatioRequest, TOrganizationResponse>> requestPart;
            chunkSize = chunkSize ?? _requestOptions.ChuckSize;

            lock (_takeAndRemoveLockObject)
            {
                var count = Math.Min(chunkSize.Value, _requests.Count());
                requestPart = _requests.Where(x => !x.IsRunning && !x.IsCompleted).Take(count).ToList();
            }
            foreach (var request in requestPart)
            {
                request.IsRunning = true;
            }

            return requestPart;
        }

        private void ExecuteRequests(Func<OrganizationServiceProxy> serviceCreator, CancellationToken cancellationToken)
        {
            var threadSw = Stopwatch.StartNew();
            IOrganizationService service;
            int reTrycount = 0;
            do
            {
                service = serviceCreator();
                if (service == null)
                {
                    Thread.Sleep(2000 * reTrycount);
                    reTrycount++;
                    OnErrorOccurs(new ErrorEventArg(new Exception($"Service is null. Service is reconnecting. RetryCount is {reTrycount}")));
                }
            } while (service == null && reTrycount < 3);
            if (service == null)
            {
                Console.WriteLine("Service is null");
                OnErrorOccurs(new ErrorEventArg(new Exception("Service has been still null. Thread is terminated")));
                OnThreadFinished(new ThreadCompletedEventArg(_tasks.Count(x => (x == null) || (x != null && /*!x.IsCompleted*/x.IsAlive)) - 1, threadSw.Elapsed, Thread.CurrentThread.Name));
                return;
            }

            List<RequestContainer<TOrganizatioRequest, TOrganizationResponse>> requests = TakeRequests(_requests);
            while (requests.Count != 0)
            {
                if (_requestOptions.IsExecuteMultipleRequest)
                {
                    ExecuteMultipleRequests(service, cancellationToken, requests);
                }
                else
                {
                    ExecuteSingleRequests(service, cancellationToken, requests);
                }
                OnChuckCompleted(new ExecutorEventArg<TOrganizatioRequest, TOrganizationResponse>(requests, _requests.Count(x => !x.IsCompleted), _requests.Count(x => x.IsRunning)));
                requests = TakeRequests(_requests);
            }
            threadSw.Stop();
            OnThreadFinished(new ThreadCompletedEventArg(_tasks.Count(x => (x == null) || (x != null && /*!x.IsCompleted*/x.IsAlive)) - 1, threadSw.Elapsed, Thread.CurrentThread.Name));
        }

        private void ExecuteSingleRequests(IOrganizationService service, CancellationToken cancellationToken, List<RequestContainer<TOrganizatioRequest, TOrganizationResponse>> requests)
        {
            for (int i = 0; i < requests.Count; i++)
            {
                try
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        break;
                    }
                    requests[i].Response = (TOrganizationResponse)service.Execute(requests[i].Request);
                    requests[i].IsCompleted = true;
                }

                catch (FaultException<OrganizationServiceFault> ex) when (ex.Detail.ErrorCode == Constants.TimeLimitExceededErrorCode)
                {
                    requests[i].AddException(ex);
                    OnRequestErrorOccurs(new RequestEventArg<TOrganizatioRequest, TOrganizationResponse>(requests[i]));
                    if (_requestOptions.IsWaitRetryAfter) Thread.Sleep((TimeSpan)ex.Detail.ErrorDetails["Retry-After"]);
                }
                catch (Exception ex)
                {
                    requests[i].AddException(ex);
                    OnRequestErrorOccurs(new RequestEventArg<TOrganizatioRequest, TOrganizationResponse>(requests[i]));
                }
                finally
                {
                    requests[i].IsRunning = false;
                }
            }
        }

        private void ExecuteMultipleRequests(IOrganizationService service, CancellationToken cancellationToken, List<RequestContainer<TOrganizatioRequest, TOrganizationResponse>> requests)
        {
            var multipleExecute = new ExecuteMultipleRequest
            {
                Settings = new ExecuteMultipleSettings
                {
                    ContinueOnError = true,
                    ReturnResponses = true,
                },
                Requests = new OrganizationRequestCollection()
            };

            for (int requestInd = 0; requestInd < requests.Count; requestInd += _executeMultipleRequestCount)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    break;
                }
                multipleExecute.Requests.Clear();
                multipleExecute.Requests.AddRange(requests.GetRange(requestInd, Math.Min(_executeMultipleRequestCount, requests.Count - requestInd)).Select(x => x.Request).ToArray());
                try
                {
                    var executeMultipleResponse = (ExecuteMultipleResponse)service.Execute(multipleExecute);
                    for (int i = 0; i < executeMultipleResponse.Responses.Count; i++)
                    {
                        var response = executeMultipleResponse.Responses[i];
                        if (response.Fault == null)
                        {
                            requests[requestInd + i].IsCompleted = true;
                            requests[requestInd + i].Response = (TOrganizationResponse)response.Response;
                        }
                        else
                        {
                            requests[requestInd + i].AddException(new FaultException<OrganizationServiceFault>(response.Fault));
                            OnRequestErrorOccurs(new RequestEventArg<TOrganizatioRequest, TOrganizationResponse>(requests[requestInd + i]));
                        }
                        requests[requestInd + i].IsRunning = false;
                    }
                }
                catch (FaultException<OrganizationServiceFault> ex) when (ex.Detail.ErrorCode == Constants.TimeLimitExceededErrorCode)
                {
                    requests.ForEach(request =>
                    {
                        request.IsRunning = false;
                        request.AddException(ex);
                        OnRequestErrorOccurs(new RequestEventArg<TOrganizatioRequest, TOrganizationResponse>(request));
                    });
                    if (_requestOptions.IsWaitRetryAfter) Thread.Sleep((TimeSpan)ex.Detail.ErrorDetails["Retry-After"]);
                }
                catch (Exception ex)
                {
                    requests.ForEach(request =>
                    {
                        request.IsRunning = false;
                        request.AddException(ex);
                        OnRequestErrorOccurs(new RequestEventArg<TOrganizatioRequest, TOrganizationResponse>(request));
                    });
                }
            }
        }

        private void CalculateStatistic()
        {
            while (_requests.Count(x => !x.IsRunning && !x.IsCompleted) != 0)
            {
                Thread.Sleep(_requestOptions.StatisticInverval);
                var completedRecordCount = _requests.Count(x => x.IsCompleted);
                var averageRecordCount = completedRecordCount / _executionSw.Elapsed.TotalSeconds;
                OnStatisticInfo(new StatisticEventArg(averageRecordCount, _requests.Count - completedRecordCount, completedRecordCount, Math.Round((double)completedRecordCount / _requests.Count, 2) * 100, _executionSw.Elapsed));
            }
        }
    }
}