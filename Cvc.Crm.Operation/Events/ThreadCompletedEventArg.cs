using System;

namespace Cvc.Crm.Operation.Events
{
    public class ThreadCompletedEventArg : EventArgs
    {
        public int RemainingThreadCount { get; }
        public string FinishedThreadName { get; }
        public TimeSpan ElapsedTime { get; }

        public ThreadCompletedEventArg(int remainingThreadCount, TimeSpan elapsedTime, string finishedThreadName = default)
        {
            RemainingThreadCount = remainingThreadCount;
            FinishedThreadName = finishedThreadName;
            ElapsedTime = elapsedTime;
        }
    }
}