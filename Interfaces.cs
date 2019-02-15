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
        bool DeliverFile(string reportName, MemoryStream memoryStream, string mimeHeader = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

    }
    public interface IServer
    {
        void CreateServer(string HTTP_IP);
  
    }

    public interface ITemplateHandler
    {
        void AddSheet(string reportName, string SheetName);
        string TemplateValidateFill(string reportName, string SheetName, string[] content);
        string TemplateValidateInit(string reportName, string SheetName, string[] header, bool createNew);
        MemoryStream ToExcel(string reportName);
    }
}

