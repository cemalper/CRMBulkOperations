using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;

namespace Cvc.Crm.Operation.Events
{
    public class CompletedEventArgs<TOrganizationRequest, TOrganizationResponse> : EventArgs where TOrganizationRequest : OrganizationRequest where TOrganizationResponse : OrganizationResponse
    {
        public List<RequestContainer<TOrganizationRequest, TOrganizationResponse>> RequestsContainers { get; }
        public TimeSpan ElapsedTime { get; }
        public double AverageRequestCount { get; }

        public CompletedEventArgs(List<RequestContainer<TOrganizationRequest, TOrganizationResponse>> requestsContainers, TimeSpan elapsedTime, double averageRequestCount)
        {
            RequestsContainers = requestsContainers;
            ElapsedTime = elapsedTime;
            AverageRequestCount = averageRequestCount;
        }
    }
}