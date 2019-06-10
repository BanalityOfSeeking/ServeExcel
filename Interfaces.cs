using System;

namespace TemplateToExcelServer.Interfaces
{

    public interface ILogger
    {
        void LogError(Exception ex, string additionalMessage);

        void Log(string additionalMessage);
    }
}