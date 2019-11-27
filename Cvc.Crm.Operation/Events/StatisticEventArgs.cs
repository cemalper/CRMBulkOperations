using System;

namespace Cvc.Crm.Operation.Events
{
    public class StatisticEventArg : EventArgs
    {
        public double AverageRequestCount { get; }
        public int FinishedRequestCount { get; }
        public int RemainingRequestCount { get; }
        public double ProgressPercentage { get; }
        public TimeSpan RunningTime { get; }

        public StatisticEventArg(double averageRequestCount, int remainingRequestCount, int finishedRequestsCount, double progressPercentage, TimeSpan runningTime)
        {
            AverageRequestCount = Math.Round(averageRequestCount, 2);
            RemainingRequestCount = remainingRequestCount;
            FinishedRequestCount = finishedRequestsCount;
            ProgressPercentage = progressPercentage;
            RunningTime = runningTime;
        }
    }
}