using System;
using System.IO;
using System.Net;

namespace ServeReports
{
    public interface ILogger
    {
        void LogError(Exception ex, string additionalMessage);

        void Log(string additionalMessage);        
    }

    public interface IHttpContextResponder
    {
        HttpListenerContext ListenerContext { get; set; }
        bool ContextPresent();
        void WriteResponse(string Message);
    }
    public interface IServer : IHttpContextResponder
    {
        void CreateServer(string HTTP_IP);
    }

    public interface ITemplateHandler
    {
        void AddSheet(string reportName, string SheetName);
        bool DeliverFile(HttpListenerContext client, string reportName, MemoryStream memoryStream, string mimeHeader = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        string TemplateValidateFill(string reportName, string SheetName, string[] content);
        string TemplateValidateInit(string reportName, string SheetName, string[] header, bool createNew);
        MemoryStream ToExcel(string reportName);
    }
}

