using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ServeReports
{

    public static class ExtensionMethods
    {
        public static bool ContentValidForLoad(this TemplateObject template) => template.ContentAdded ? template.ContentArray.Length % template.FormatLength == 0 ? true : false : false;

        public static void ToFile(this MemoryStream ms, string FileName)
        {
            FileStream fs = File.Create(FileName);
            ms.CopyTo(fs);

            fs.Flush();
            ms.Flush();

            ms.Close();
            fs.Close();
        }
    }
    public class TemplateObjectHandler : ITemplateObjectHandler
    {

        public ILogger log;
        public TemplateContainer Container { get; private set; }

        public TemplateObjectHandler(TemplateContainer container, ILogger logger)
        {
            Container = container;
            log = logger;

        }
        public TemplateObject GetTemplateReportSheet(string ReportName, string SheetName)
        {
            return GetTemplateReportSheets(ReportName).Where(x => x.SheetName == SheetName).FirstOrDefault();
        }

        public IEnumerable<TemplateObject> GetTemplateReportSheets(string ReportName)
        {
            return GetTemplateReports().Where(x => x.NameOfReport == ReportName);
        }
        private IEnumerable<TemplateObject> GetTemplateReports()
        {
            return (from tObject in Container.GetTObject()
                    select tObject);
        }
        
        internal bool TemplateObjectUpdate(string ReportName, string SheetName, string[] header)
        {

            TemplateObject TempObj = GetTemplateReportSheet(ReportName,SheetName);

            TempObj.Format = header;

            return true;
        }
        public bool TemplateObjectCreate(string ReportName, string SheetName, string[] header)
        {
            if (TemplateObjectUpdate(ReportName, SheetName, header))
            {
                return true;
            }
            TemplateObject tempObj = new TemplateObject
            {
                NameOfReport = ReportName,

                SheetName = SheetName
            };
            tempObj.Format = header;

            Container.AddTObject(tempObj);
            return true;
        }
        private bool TemplateObjectCreateOrUpdate(string reportName, string SheetName, string[] header, bool createNew)
        {
            if (createNew)
            {
                if (TemplateObjectCreate(reportName, SheetName, header))
                {
                    return true;
                }
                return false;
            }
            else
            {
                if (TemplateObjectUpdate(reportName, SheetName, header))
                {
                    return true;
                }
                return false;
            }
        }
        public bool TemplateFill(string reportName, string SheetName, string[] Content)
        {
            List<string> ContentList = Content.ToList();
            
            TemplateObject TempObj = Container.GetTObject().Find(x => x.SheetName == SheetName);

            if (TempObj == null) return false;

            if (Content.Length % TempObj.FormatLength == 0)
            {
                for (int r = 0; r < TempObj.NumberOfRowsToAdd; r++)
                {
                    var ListRow = Content.Take(TempObj.FormatLength.Value).ToList();
                    for (int c = 0; c < ListRow.Count(); c++)
                    {
                        TempObj.Content[r,c] = ListRow[c];
                    }
                }
                TempObj.ContentArray = Content;

                return true;
            }
            return false;

        }
        public void TemplateObjectToDataTable(string ReportName, string SheetName)
        {

            DataTable dt = new DataTable(SheetName)
            {
                Locale = CultureInfo.CurrentCulture
            };

            TemplateObject TempObj = Container.GetTObject().Find(x => x.SheetName == SheetName);

            for (int i = 1; i < TempObj.Format.Length + 1; i++)
            {
                dt.Columns.Add(TempObj.Format[i - 1]);
            }

            int NumberOfRowsToAdd = TempObj.Content.Length / TempObj.Format.Length;

            for (int rows = 0; rows < NumberOfRowsToAdd; rows++)
            {
                string[] array = new string[TempObj.Format.Length];
                for (int columns = 0; columns < TempObj.Format.Length; columns++)
                {
                    array[columns] = TempObj.Content[rows, columns];
                }
                object[] dr = dt.NewRow().ItemArray = array;
                dt.Rows.Add(dr);
            }
            TempObj.DataTable = dt;
        }
        public MemoryStream ToExcel(string ReportName)
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

                foreach (var TempObj in Container.GetTObject().Where(x => x.NameOfReport == ReportName))
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

                    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = TempObj.DataTable.TableName };
                    sheets.Append(sheet);

                    Row headerRow = new Row();

                    List<string> columns = new List<string>();

                    foreach (DataColumn column in TempObj.DataTable.Columns)
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
                    foreach (DataRow dsrow in TempObj.DataTable.Rows)
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
    }
}

