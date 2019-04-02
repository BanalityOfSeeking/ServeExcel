using System;
using Template.Interfaces;

namespace Template.Logger
{
    public class Logger : ILogger
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