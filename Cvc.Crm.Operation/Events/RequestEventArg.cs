using Microsoft.Xrm.Sdk;
using System;

namespace Cvc.Crm.Operation.Events
{
    public class RequestEventArg<TOrganizationRequest, TOrganizationResponse> : EventArgs where TOrganizationRequest : OrganizationRequest where TOrganizationResponse : OrganizationResponse
    {
        public RequestContainer<TOrganizationRequest, TOrganizationResponse> RequestContainer { get; }

        public RequestEventArg(RequestContainer<TOrganizationRequest, TOrganizationResponse> requestContainer)
        {
            RequestContainer = requestContainer;
        }
    }
}