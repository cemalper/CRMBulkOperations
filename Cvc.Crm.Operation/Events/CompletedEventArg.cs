using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;

namespace Cvc.Crm.Operation.Events
{
    public class CompletedEventArgs<TOrganizationRequest, TOrganizationResponse> : EventArgs where TOrganizationRequest : OrganizationRequest where TOrganizationResponse : OrganizationResponse
    {
        public List<RequestContainer<TOrganizationRequest, TOrganizationResponse>> RequestsContainers { get; }
        public TimeSpan ElapsedTime { get; }

        public CompletedEventArgs(List<RequestContainer<TOrganizationRequest, TOrganizationResponse>> requestsContainers, TimeSpan elapsedTime)
        {
            RequestsContainers = requestsContainers;
            ElapsedTime = elapsedTime;
        }
    }
}