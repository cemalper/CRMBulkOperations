using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;

namespace Cvc.Crm.Operation.Events
{
    public class ExecutorEventArg<TOrganizationRequest, TOrganizationResponse> : EventArgs where TOrganizationRequest : OrganizationRequest where TOrganizationResponse : OrganizationResponse
    {
        public List<RequestContainer<TOrganizationRequest, TOrganizationResponse>> RequestsContainers { get; }
        public int RemainingRequestCount { get; }

        public int RunningRequestCount { get; }

        public ExecutorEventArg(List<RequestContainer<TOrganizationRequest, TOrganizationResponse>> requestContainers, int remainingRequestCount, int runningRequestCount)
        {
            RequestsContainers = requestContainers;
            RemainingRequestCount = remainingRequestCount;
            RunningRequestCount = runningRequestCount;
        }
    }
}