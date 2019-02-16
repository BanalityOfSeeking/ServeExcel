using System;
using System.Collections.Generic;
using System.Data;
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
    public interface ITemplateObject
    {
        string NameOfReport { get; set; }
        string SheetName { get; set; }
        string[] Format { get; set; }
        string[,] Content { get; set; }
        string[] ContentArray { get; set; }
        int? FormatLength { get; }
        int? ContentArrayLength { get; }
        DataTable DataTable { get; set; }
    }
    public interface ITemplateContainer
    {
        List<TemplateObject> GetTObject();
        void AddTObject(TemplateObject templateObject);
    }
    public interface ITemplateObjectHandler
    {
        TemplateContainer Container { get; }

        TemplateObject GetTemplateReportSheet(string ReportName, string SheetName);
        IEnumerable<TemplateObject> GetTemplateReportSheets(string ReportName);
        bool TemplateFill(string reportName, string SheetName, string[] content);
        bool TemplateObjectCreate(string ReportName, string SheetName, string[] header);
        void TemplateObjectToDataTable(string ReportName, string SheetName);
        MemoryStream ToExcel(string ReportName);
    }
}

