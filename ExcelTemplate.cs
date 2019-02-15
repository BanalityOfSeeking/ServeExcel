using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Formatters.Binary;

namespace ServeReports
{

    public class TemplateHandler : ITemplateHandler
    {
        private readonly ILogger _logger;
        public TemplateHandler(ILogger logger)
        {
            _logger = logger;
        }
        internal Dictionary<(string, string), byte[]> InMemContainers = new Dictionary<(string, string), byte[]>();

        internal Dictionary<(string, string), DataTable> sheetTables = new Dictionary<(string, string), DataTable>();

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
                if (TemplateCreateOrUpdate(reportName, SheetName, header, createNew))
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
        internal bool TemplateUpdate(string reportName, string SheetName, string[] header)
        {
            BinaryFormatter binFormatter = new BinaryFormatter();
            if (InMemContainers.ContainsKey((reportName, SheetName)))
            {

                using (MemoryStream ms = new MemoryStream(InMemContainers[(reportName, SheetName)]))
                {

                    TemplateContainer Template = binFormatter.Deserialize(ms) as TemplateContainer;

                    TemplateObject TempObj = Template.TObject.Find(x => x.SheetName == SheetName);

                    TempObj.Format = new string[header.Length];

                    header.CopyTo(TempObj.Format, 0);

                    if (Template.ContentLength > 0 && Template.ContentLength % Template.FormatLength == 0)
                    {
                        if (TemplateFill(reportName, SheetName, Template.ContentArray))
                        {

                            binFormatter.Serialize(ms, Template);

                            return true;

                        }
                    }
                }
            }
            return false;
        }
        internal bool TemplateCreate(string reportName, string SheetName, string[] header)
        {
            int cc = InMemContainers.Count();
            if (InMemContainers.ContainsKey((reportName, SheetName)))
            {
                if (TemplateUpdate(reportName, SheetName, header))
                {
                    return true;
                }
                return false;
            }

            BinaryFormatter binFormatter = new BinaryFormatter();

            var Template = new TemplateContainer();

            TemplateObject tempObj = new TemplateObject
            {
                NameOfReport = reportName,

                SheetName = SheetName
            };

            byte[] MemoryBuffer = new byte[100 * 1024];

            using (MemoryStream ms = new MemoryStream(MemoryBuffer, true))
            {

                tempObj.Format = new string[Template.FormatLength];

                header.CopyTo(tempObj.Format, 0);

                Template.TObject.Add(tempObj);

                Template.FormatLength = header.Count();

                binFormatter.Serialize(ms, Template);

                InMemContainers.Add((reportName, SheetName), ms.ToArray());
                if (InMemContainers.Count() > cc)
                {
                    return true;
                }
                return false;

            }

        }
        private bool TemplateCreateOrUpdate(string reportName, string SheetName, string[] header, bool createNew)
        {
            if (createNew)
            {
                if (TemplateCreate(reportName, SheetName, header))
                {
                    return true;
                }
                return false;
            }
            else
            {
                if (TemplateUpdate(reportName, SheetName, header))
                {
                    return true;
                }
                return false;
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
                if (TemplateFill(reportName, SheetName, content))
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

        private bool TemplateFill(string reportName, string SheetName, string[] content)
        {

            string reportString = reportName;

            BinaryFormatter binFormatter = new BinaryFormatter();

            if (InMemContainers.ContainsKey((reportName, SheetName)))
            {
                using (MemoryStream ms = new MemoryStream(InMemContainers[(reportName, SheetName)]))
                {
                    TemplateContainer Template = binFormatter.Deserialize(ms) as TemplateContainer;

                    if (content.Length % Template.FormatLength == 0)
                    {
                        int NumberOfRowsToAdd = content.Length / Template.FormatLength;
                        int row = 0;
                        int Column = 0, pColumn = 0;
                        TemplateObject tempObject = Template.TObject.Find(x => x.SheetName == SheetName);
                        tempObject.Content = new string[NumberOfRowsToAdd, Template.FormatLength];
                        do
                        {
                            tempObject.Content[row, pColumn] = content[Column];
                            Column = Column + 1;
                            pColumn = pColumn + 1;
                            if (pColumn == Template.FormatLength)
                            {
                                row = row + 1;
                                pColumn = 0;
                            }
                        } while (Column < (Template.FormatLength * NumberOfRowsToAdd));

                        Template.ContentArray = content;

                        Template.ContentLength = content.Length;

                        ms.Seek(0, SeekOrigin.Begin);

                        binFormatter.Serialize(ms, Template);

                        return true;
                    }
                }
            }
            return false;

        }

        public void AddSheet(string reportName, string SheetName)
        {

            DataTable dt = new DataTable(SheetName)
            {
                Locale = CultureInfo.CurrentCulture
            };

            BinaryFormatter binFormatter = new BinaryFormatter();

            if (InMemContainers.ContainsKey((reportName, SheetName)))
            {

                using (MemoryStream ms = new MemoryStream(InMemContainers[(reportName, SheetName)]))
                {

                    var Template = binFormatter.Deserialize(ms) as TemplateContainer;

                    TemplateObject tempObject = Template.TObject.Find(x => x.SheetName == SheetName);

                    for (int i = 1; i < tempObject.Format.Length + 1; i++)
                    {
                        dt.Columns.Add(tempObject.Format[i - 1]);
                    }

                    int NumberOfRowsToAdd = tempObject.Content.Length / tempObject.Format.Length;

                    for (int rows = 0; rows < NumberOfRowsToAdd; rows++)
                    {
                        string[] array = new string[tempObject.Format.Length];
                        for (int columns = 0; columns < tempObject.Format.Length; columns++)
                        {
                            array[columns] = tempObject.Content[rows, columns];
                        }
                        object[] dr = dt.NewRow().ItemArray = array;
                        dt.Rows.Add(dr);
                    }
                    if (!sheetTables.ContainsKey((reportName, SheetName)))
                    {
                        sheetTables.Add((reportName, SheetName), dt);
                    }
                    return;
                }
            }
            else
            {

                return;
            }
        }

        public MemoryStream ToExcel(string reportName)
        {

            byte[] FileBuffer = new byte[1024 * 100];
            MemoryStream memoryStream = new MemoryStream(FileBuffer, true);

            using (SpreadsheetDocument workbook = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook, true))
            {
                workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new Workbook
                {
                    Sheets = new Sheets()
                };

                uint sheetId = 1;

                foreach (var table in sheetTables.Where(x => x.Key.Item1 == reportName))
                {
                    WorksheetPart sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();

                    SheetData sheetData = new SheetData();

                    sheetPart.Worksheet = new Worksheet(sheetData);

                    Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();

                    string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    if (sheets.Elements<Sheet>().Count() > 0)
                    {
                        sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.Value.TableName };
                    sheets.Append(sheet);

                    Row headerRow = new Row();

                    List<string> columns = new List<string>();

                    foreach (DataColumn column in table.Value.Columns)
                    {
                        columns.Add(column.ColumnName);

                        Cell cell = new Cell
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(column.ColumnName)
                        };

                        headerRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(headerRow);
                    foreach (DataRow dsrow in table.Value.Rows)
                    {
                        Row newRow = new Row();
                        foreach (string col in columns)
                        {
                            Cell cell = new Cell
                            {
                                DataType = CellValues.String,
                                CellValue = new CellValue(dsrow[col].ToString())
                            };
                            newRow.AppendChild(cell);
                        }
                        sheetData.AppendChild(newRow);
                    }
                }
            }
            return memoryStream;

        }
        public bool DeliverFile(HttpListenerContext client, string reportName, MemoryStream memoryStream, string mimeHeader = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        {
            using (HttpListenerResponse response = client.Response)
            {
                response.StatusCode = 200;
                response.StatusDescription = "OK";
                response.ContentType = mimeHeader;
                response.ContentLength64 = memoryStream.Length;
                response.AddHeader("Content-disposition", "attachment; filename=" + reportName);
                response.SendChunked = false;
                using (BinaryWriter bw = new BinaryWriter(response.OutputStream))
                {
                    byte[] bContent = memoryStream.ToArray();
                    bw.Write(bContent, 0, bContent.Length);
                    bw.Flush();
                    bw.Close();
                    return true;
                }
            }
        }
    }
}

