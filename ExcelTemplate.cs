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
using System.Text;

namespace ServeReports
{
    public class ExcelTemplate
    {
        [Serializable]
        public struct TemplateObject
        {
            public string NameOfReport { get; set; }
            public string[] Format { get; set; }
            public string[,] Content { get; set; }
        }
        [Serializable]
        public class TemplateContainer
        {
            public TemplateObject TObject = new TemplateObject();

            public string[] ContentArray { get; set; }
            public int FormatLength { get; set; }
            public int ContentLength { get; set; }
        }

        public static bool TemplateInit(string reportName, string[] header, bool createNew, HttpListenerContext socket)
        {
            if (TemplateInit(reportName, header, createNew))
            {
                socket.Response.OutputStream.Write(Encoding.UTF8.GetBytes("Successfully Initialized " + reportName), 0, ("Successfully Initialized " + reportName).Length);
                return true;
            }
            socket.Response.OutputStream.Write(Encoding.UTF8.GetBytes("Failed Initialization " + reportName), 0, ("Failed Initialization " + reportName).Length);
            return false;
        }

        //initializes	Excel	Template
        //<Parameter>	ReportName:	Names	the	Sheet	and	Excel
        //<Parameter>	Header	:	array	of	values
        //<Parameter>	CreateNew	:	boolean	indicating	if	the	template	needs	to	be	created	or	updated.
        public static bool TemplateInit(string reportName, string[] header, bool createNew)
        {
            string reportString = reportName;
            BinaryFormatter binFormatter = new BinaryFormatter();
            if (header.Length != 0)
            {
                //Handle	Creation	of	new	template
                if (createNew)
                {
                    try
                    {
                        //build	and	store	template
                        TemplateContainer Template = new TemplateContainer();
                        //name	the	report
                        Template.TObject.NameOfReport = reportName;
                        //StoreLength
                        Template.FormatLength = header.Length;
                        //initialize	the	format	string
                        Template.TObject.Format = new string[Template.FormatLength];
                        //fill	out	the	Format
                        header.CopyTo(Template.TObject.Format, 0);
                        //serialize	Template	to	remember	the	reports	we	have	setup.
                        using (FileStream fs = new FileStream(Directory.GetCurrentDirectory() + "\\Configs\\" + reportString, FileMode.Create, FileAccess.Write, FileShare.None))
                        {
                            binFormatter.Serialize(fs, Template);
                        }
                        binFormatter = null;
                        return true;
                    }
                    //write	all	exceptions	to	console	window	on	server
                    catch
                    {
                        throw;
                    }
                    //clean	up
                    finally
                    {
                        binFormatter = null;
                    }
                }
                //UPDATE	Template
                else
                {
                    try
                    {
                        TemplateContainer Template = new TemplateContainer();
                        //check	if	report	config	exists
                        if (File.Exists(Directory.GetCurrentDirectory() + "\\Configs\\" + reportString))
                        {
                            //Deserialize	TemplateObject
                            using (FileStream fs = new FileStream(Directory.GetCurrentDirectory() + "\\Configs\\" + reportString, FileMode.Open, FileAccess.Read, FileShare.None))
                            {
                                Template = (TemplateContainer)binFormatter.Deserialize(fs);

                                //write	out	the	header
                                Template.TObject.Format = new string[header.Length];
                                header.CopyTo(Template.TObject.Format, 0);
                                //realign	content	if	possible, else content is removed.***!!!****
                                if (Template.ContentLength > 0 && Template.ContentLength % Template.FormatLength == 0)
                                {
                                    TemplateFill(reportName, Template.ContentArray);
                                }
                            }
                            using (FileStream fs = new FileStream(Directory.GetCurrentDirectory() + "\\Configs\\" + reportString, FileMode.Open, FileAccess.Write, FileShare.None))
                            {
                                binFormatter.Serialize(fs, Template);
                            }
                            return true;
                        }
                        else
                        {
                            //Console.Write("Configuration does not exist, can not update what does not exist.");
                            return false;
                        }
                    }
                    catch
                    {
                        throw;
                    }
                    finally
                    {
                        binFormatter = null;
                    }
                }
            }
            return false;
        }

        public static bool TemplateFill(string reportName, string[] content, HttpListenerContext socket)
        {
            if (TemplateFill(reportName, content))
            {
                socket.Response.OutputStream.Write(Encoding.UTF8.GetBytes(reportName + " successfully filled"), 0, (reportName + " successfully filled").Length);
                return true;
            }
            socket.Response.OutputStream.Write(Encoding.UTF8.GetBytes(reportName + " fill failed"), 0, (reportName + " fill failed").Length);
            return false;
        }

        //Fill	in	Template	content
        //<Parameter>	ReportName:	Names	the	Sheet	and	Excel
        //<Parameter>	Content	:	array	of	values

        public static bool TemplateFill(string reportName, string[] content)
        {
            if (content == null)
            {
                throw new ArgumentNullException(nameof(content));
            }

            string reportString = reportName;
            BinaryFormatter binFormatter = new BinaryFormatter();
            TemplateContainer Template = new TemplateContainer();
            try
            {
                //check	if	report	config	exists
                if (File.Exists(Directory.GetCurrentDirectory() + "\\Configs\\" + reportString))
                {
                    using (FileStream fs = new FileStream(Directory.GetCurrentDirectory() + "\\Configs\\" + reportString, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        Template = (TemplateContainer)binFormatter.Deserialize(fs);
                    }
                    if (content.Length % Template.FormatLength == 0)
                    {
                        int NumberOfRowsToAdd = content.Length / Template.FormatLength;
                        int row = 0;
                        int Column = 0, pColumn = 0;
                        Template.TObject.Content = new string[NumberOfRowsToAdd, Template.FormatLength];
                        do
                        {
                            Template.TObject.Content[row, pColumn] = content[Column];
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
                        using (FileStream fe = new FileStream(Directory.GetCurrentDirectory() + "\\Configs\\" + reportString, FileMode.Open, FileAccess.Write, FileShare.None))
                        {
                            binFormatter.Serialize(fe, Template);
                        }
                        binFormatter = null;
                        return true;
                    }
                }
                return false;
            }
            catch
            {
                throw;
            }
            finally
            {
                binFormatter = null;
            }
        }

        //Build	the	Excel	File	from	TemplateObject

        public static void AddSheet(string reportName, ref DataSet ds)
        {
            string reportString = reportName;
            DataTable dt = new DataTable(reportName)
            {
                Locale = CultureInfo.CurrentCulture
            };
            //init	serializer
            BinaryFormatter binFormatter = new BinaryFormatter();
            try
            {
                //check	if	report	config	exists
                if (File.Exists(Directory.GetCurrentDirectory() + "\\Configs\\" + reportString))
                {
                    //Deserialize	it
                    using (FileStream fs = new FileStream(Directory.GetCurrentDirectory() + "\\Configs\\" + reportString, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        TemplateContainer template = (TemplateContainer)binFormatter.Deserialize(fs);
                        //write	out	the	format
                        for (int i = 1; i < template.TObject.Format.Length + 1; i++)
                        {
                            dt.Columns.Add(template.TObject.Format[i - 1]);
                        }
                        //get	the	numer	of	rows	to	add
                        int NumberOfRowsToAdd = template.TObject.Content.Length / template.TObject.Format.Length;
                        //get	the	working	row

                        for (int rows = 0; rows < NumberOfRowsToAdd; rows++)
                        {
                            string[] array = new string[template.TObject.Format.Length];
                            for (int columns = 0; columns < template.TObject.Format.Length; columns++)
                            {
                                array[columns] = template.TObject.Content[rows, columns];
                            }
                            object[] dr = dt.NewRow().ItemArray = array;
                            dt.Rows.Add(dr);
                        }
                        ds.Tables.Add(dt);
                        binFormatter = null;
                        return;
                        //handle	error
                    }
                }
                else
                {
                    //log.Error("No Configuration file setup for this action.");
                    return;
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                binFormatter = null;
            }
        }

        public static void DataSetToExcel(DataSet ds, string destination)
        {
            try
            {
                using (SpreadsheetDocument workbook = SpreadsheetDocument.Create(Directory.GetCurrentDirectory() + "\\Worksheets\\" + destination + ".xlsx", SpreadsheetDocumentType.Workbook, true))
                {
                    workbook.AddWorkbookPart();
                    workbook.WorkbookPart.Workbook = new Workbook
                    {
                        Sheets = new Sheets()
                    };

                    uint sheetId = 1;

                    foreach (DataTable table in ds.Tables)
                    {
                        //  workbook.WorkbookPart.Workbook.ExcelNamedRange(table.TableName, table.TableName, "A", "1", "I", "1");
                        WorksheetPart sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                        SheetData sheetData = new SheetData();
                        sheetPart.Worksheet = new Worksheet(sheetData);

                        Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                        string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                        if (sheets.Elements<Sheet>().Count() > 0)
                        {
                            sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                        }

                        Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                        sheets.Append(sheet);

                        Row headerRow = new Row();

                        List<string> columns = new List<string>();

                        foreach (DataColumn column in table.Columns)
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
                        foreach (DataRow dsrow in table.Rows)
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
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static void DeliverFile(HttpListenerContext client, string fileName, int contentLength, string mimeHeader = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", int statusCode = 200)
        {
            using (HttpListenerResponse response = client.Response)
            {
                try
                {
                    string prevDirectory = Directory.GetCurrentDirectory();
                    Directory.SetCurrentDirectory(prevDirectory + "\\Worksheets\\");
                    response.StatusCode = statusCode;
                    response.StatusDescription = "OK";
                    response.ContentType = mimeHeader;
                    response.ContentLength64 = contentLength;
                    response.AddHeader("Content-disposition", "attachment; filename=" + fileName);
                    response.SendChunked = false;
                    using (BinaryWriter bw = new BinaryWriter(response.OutputStream))
                    {

                        byte[] bContent = File.ReadAllBytes(fileName);
                        bw.Write(bContent, 0, bContent.Length);
                        bw.Flush(); //seems to have no effect
                        bw.Close();
                    }
                    Directory.SetCurrentDirectory(prevDirectory);
                    //Console.WriteLine("Total Bytes : " + ContentLength.ToString());
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }
    }
}

