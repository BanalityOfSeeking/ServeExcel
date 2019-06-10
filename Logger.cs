using System;
using TemplateToExcelServer.Interfaces;

namespace TemplateToExcelServer.Logger
{
    public class Logger : ILogger
    {

        public virtual void Log(string message)
        {
            Console.WriteLine(message);
        }
        public virtual void LogError(Exception ex, string additionalMessage)
        {
            Console.WriteLine($"{additionalMessage}: {ex}");
        }
    }
}