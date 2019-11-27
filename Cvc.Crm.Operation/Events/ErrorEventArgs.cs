using System;

namespace Cvc.Crm.Operation.Events
{
    public class ErrorEventArg : EventArgs
    {
        public Exception Exception { get; }

        public ErrorEventArg(Exception exception)
        {
            Exception = exception;
        }
    }
}