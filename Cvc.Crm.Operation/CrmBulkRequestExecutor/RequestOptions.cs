namespace Cvc.Crm.Operation
{
    public class RequestOptions
    {
        public bool IsExecuteMultipleRequest { get; set; } = true;
        public int ChuckSize { get; set; } = 40;
        public int ParalelSize { get; set; } = 16;
        public int RetryCount { get; set; } = 3;
        public bool IsWaitRetryAfter { get; set; } = false;

        [System.Timers.TimersDescription("The number of milliseconds between timer events.")]
        public int StatisticInverval { get; set; } = 4000;
    }
}