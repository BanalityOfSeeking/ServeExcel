using System;

namespace ServeReports
{
    public class ConsoleLogger : ILogger
    {

        public void LogError(Exception ex, string additionalMessage)
        {
            Console.WriteLine("{0}: {1}", additionalMessage, ex);
        }

        public void Log(string message)
        {
            Console.WriteLine(message);
        }
    }
}

