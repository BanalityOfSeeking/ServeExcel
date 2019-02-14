using System;
using System.Net;

namespace ServeReports
{
    public interface ILogger
    {
        void LogError(Exception ex, string additionalMessage);

        void Log(string additionalMessage);        
    }
    public interface ISocketLogger 
    {
        ConsoleLogger ServerLog { get; set; }
        void ClientError(HttpListenerContext client, Exception ex, string additionalMessage);
        void ClientLog(HttpListenerContext client, string additionalMessage);

    }
}

