using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;

namespace ServeReports
{

    public class Temphandler
    {
        private readonly SocketLogger _logger;
        public Temphandler(SocketLogger logger)
        {
            _logger = logger;
        }
        internal Dictionary<(string, string), byte[]> InMemContainers = new Dictionary<(string, string), byte[]>();

        internal Dictionary<(string, string), DataTable> sheetTables = new Dictionary<(string, string), DataTable>();

        public bool TemplateInit(string reportName, string SheetName, string[] header, bool createNew, HttpListenerContext socket)
        {
            if (string.IsNullOrEmpty(reportName) | string.IsNullOrEmpty(SheetName))
            {
                _logger.ClientLog(socket, "Failed report name or sheet name cannot be null");
                return false;
            }

            if (header.Count() == 0 & header != null)
            {
                _logger.ClientLog(socket,"Failed headers cannot be null");
                return false;
            }

            if (TemplateInit(reportName, SheetName, header, createNew))
            {
                _logger.ClientLog(socket,"Successfully Initialized " + reportName + " with Sheet " + SheetName);
                return true;
            }
            return false;


        }
        private bool TemplateInit(string reportName, string SheetName, string[] header, bool createNew)
        {
            BinaryFormatter binFormatter = new BinaryFormatter();

            byte[] MemoryBuffer = new byte[10 * 1024];

            MemoryStream ms = new MemoryStream(MemoryBuffer, true);

            if (header.Length != 0)
            {

                if (createNew)
                {

                    var Template = new TemplateContainer();

                    TemplateObject tempObj = new TemplateObject
                    {
                        NameOfReport = reportName,

                        SheetName = SheetName
                    };

                    Template.TObject.Add(tempObj);

                    Template.FormatLength = header.Count();

                    tempObj.Format = new string[Template.FormatLength];

                    header.CopyTo(tempObj.Format, 0);

                    binFormatter.Serialize(ms, Template);

                    InMemContainers.Add((reportName, SheetName), ms.ToArray());

                    return true;
                }
            }
            else
            {
                if (InMemContainers.ContainsKey((reportName, SheetName)))
                {

                    using (ms = new MemoryStream(InMemContainers[(reportName, SheetName)]))
                    {

                        TemplateContainer Template = binFormatter.Deserialize(ms) as TemplateContainer;

                        TemplateObject TempObj = Template.TObject.Find(x => x.SheetName == SheetName);

                        TempObj.Format = new string[header.Length];

                        header.CopyTo(TempObj.Format, 0);

                        if (Template.ContentLength > 0 && Template.ContentLength % Template.FormatLength == 0)
                        {
                            if (TemplateFill(reportName, SheetName, Template.ContentArray))
                            {
                                ms.Seek(0, SeekOrigin.Begin);

                                binFormatter.Serialize(ms, Template);

                                return true;

                            }
                        }

                    }
                }
            }
            return false;
        }


        public bool TemplateFill(string reportName, string SheetName, string[] content, HttpListenerContext socket)
        {
            if (string.IsNullOrEmpty(reportName) | string.IsNullOrEmpty(SheetName))
            {
                return false;
            }
            if (content.Count() == 0 & content != null)
            {
                return false;
            }
            if (TemplateFill(reportName, SheetName, content))
            {
                socket.Response.OutputStream.Write(Encoding.UTF8.GetBytes(reportName + " successfully filled"), 0, (reportName + " successfully filled").Length);
                return true;
            }
            socket.Response.OutputStream.Write(Encoding.UTF8.GetBytes(reportName + " fill failed"), 0, (reportName + " fill failed").Length);
            return false;
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
                                CellValue = new CellValue(dsrow[col].ToString())    //
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

