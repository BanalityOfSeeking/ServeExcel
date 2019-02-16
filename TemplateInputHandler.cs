using System;

namespace ServeReports
{

    public class TemplateInputHandler 
    {
        private readonly ILogger _logger;
        public TemplateInputHandler(TemplateObjectHandler TemplateHandler, ILogger logger)
        {
            _logger = logger;
            ObjectHandler = TemplateHandler;
        }
        TemplateObjectHandler ObjectHandler { get; set; }
        public string TemplateValidateInit(string reportName, string SheetName, string[] header, bool createNew)
        {
            if (string.IsNullOrEmpty(reportName))
            {
                _logger.Log("Failed report name cannot be null");
                return ("Failed report name cannot be null");
            }
            if (string.IsNullOrEmpty(SheetName))
            {
                _logger.Log("Failed SheetName cannot be null");
                return ("Failed SheetName cannot be null");
            }
            if (header.Length == 0 | header == null)
            {
                _logger.Log("Initialization Failed, headers cannot be null");
                return ("Initialization Failed, headers cannot be null");
            }
            try
            {
                if (ObjectHandler.TemplateObjectCreate(reportName, SheetName, header))
                {                    
                    _logger.Log("Successfully Initialized " + reportName + " with Sheet " + SheetName);
                    return ("Successfully Initialized " + reportName + " with Sheet " + SheetName);
                }
               
                _logger.Log("Failed Initializing Template " + reportName + " with Sheet " + SheetName);
                return ("Failed Initializing Template " + reportName + " with Sheet " + SheetName);
                
            }
            catch(Exception ex)
            {
                
                _logger.LogError(ex, "Failed Initializing Template " + reportName + " with Sheet " + SheetName);
                return ("Failed Initializing Template " + reportName + " with Sheet " + SheetName);

            }
        }

        public string TemplateValidateFill(string reportName, string SheetName, string[] content)
        {

            if (string.IsNullOrEmpty(reportName))
            {
                _logger.Log("Failed report name cannot be null");
                return ("Failed report name cannot be null");
            }
            if (string.IsNullOrEmpty(SheetName))
            {
                _logger.Log("Failed SheetName cannot be null");
                return ("Failed SheetName cannot be null");
            }
            if (content.Length == 0 | content == null)
            {
                _logger.Log("content parameter cannot be blank");
                return ("content parameter cannot be blank or null");
            }
            try
            {
                if (ObjectHandler.TemplateFill(reportName, SheetName, content))
                {
                    _logger.Log(reportName + " " + SheetName + " successfully filled");
                    return (reportName + " " + SheetName + " successfully filled");
                }
                _logger.Log(reportName + " " + SheetName + " fill failed");
                return (reportName + " " + SheetName + " fill failed");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, reportName + " fill failed");
                return (reportName + " fill failed");
            }
        }
    }
}

