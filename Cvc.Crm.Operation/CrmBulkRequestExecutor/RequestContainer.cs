using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.ServiceModel;
namespace Cvc.Crm.Operation
{
    public class RequestContainer<TOrganizaitonRequest, TOrganizationResponse> where TOrganizaitonRequest : OrganizationRequest where TOrganizationResponse : OrganizationResponse
    {
        public TOrganizaitonRequest Request { get; private set; }
        public TOrganizationResponse Response { get; set; }
        public int TryCount { get; set; } = 0;
        public List<Exception> Exceptions { get; }
        public bool IsRunning { get; set; }
        public bool IsCompleted { get; set; }
        public bool IsFailed { get; set; }
        public int TryCountLimit { get; }

        public RequestContainer(TOrganizaitonRequest request, int tryCountLimit)
        {
            Request = request;
            TryCountLimit = tryCountLimit;
            Exceptions = new List<Exception>();
        }

        public void AddException(Exception exception)
        {
            Exceptions.Add(exception);
            if ((exception
                 is FaultException && ((FaultException<OrganizationServiceFault>)exception).Detail.ErrorCode == Constants.TimeLimitExceededErrorCode) == false)
            {
                TryCount++;
            }
            if (TryCount == TryCountLimit)
            {
                IsCompleted = true;
                IsFailed = true;
            }
        }
    }
}