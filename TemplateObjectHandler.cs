using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Template.HttpResponder;
using Template.Interfaces;
using Template.Template;
using Template.TemplateContainer;

namespace Template.TemplateHandler
{
    public partial class TemplateObjectHandler
    {
        public enum CommandType
        {
            CreateTemplate,
            ModifyTemplate,
            AppendContent,
            AddContent
        }

        public readonly ILogger _logger;
        public TemplateContainer<TemplateObject> Container { get; private set; }

        public TemplateObjectHandler(TemplateContainer<TemplateObject> container, ILogger logger)
        {
            Container = container;
            _logger = logger;
        }

        public bool TemplateObjectCreateUpdateAppendOverwrite(ReadOnlySpan<byte> ReportName, ReadOnlySpan<byte> SheetName, CommandType CommandType, string[] header, string[] Content)
        {
            switch (CommandType)
            {
                case CommandType.CreateTemplate://create
                    {
                        if (header == null)
                        {
                            return false;
                        }
                        TemplateObject TempObj = new TemplateObject()
                            .SetNameOfReport(ReportName.ToArray())
                            .SetSheetName(SheetName.ToArray())
                            .SetFormat(header);

                        Container.AddTemplate(TempObj);
                        break;
                    }
                case CommandType.ModifyTemplate://update
                    {
                        TemplateObject TempObj = Container.GetTemplateSheet(ReportName, SheetName);
                        if (TempObj != null)
                        {
                            TempObj.SetFormat(header);
                            break;
                        }
                        return false;
                    }
                case CommandType.AppendContent://append content
                    {
                        int TotalCount = 0;
                        TemplateObject TempObj = Container.GetTemplateSheet(ReportName, SheetName);
                        if (TempObj != null)
                        {
                            if (TempObj.NumberOfRowsToAdd.HasValue)
                            {
                                List<string> lca = TempObj.GetContentArray().ToList();
                                lca.AddRange(Content);
                                TempObj.SetContentArray(lca.ToArray());

                                TempObj.SetContent(new string[TempObj.NumberOfRowsToAdd.Value, TempObj.GetFormat().Length]);
                                for (int ri = 0; ri < TempObj.NumberOfRowsToAdd; ri++)
                                {
                                    for (int ci = 0; ci < TempObj.GetFormat().Length; ci++)
                                    {
                                        TempObj.GetContent()[ri, ci] = TempObj.GetContentArray()[TotalCount];
                                        TotalCount += 1;
                                    }
                                }
                                if (TempObj.ContentValidForLoad)
                                {
                                    break;
                                }
                            }
                        }
                        return false;
                    }
                case CommandType.AddContent:
                    {
                        int TotalCount = 0;
                        TemplateObject TempObj = Container.GetTemplateSheet(ReportName, SheetName);
                        TempObj.SetContentArray(Content);
                        TempObj.SetContent(new string[TempObj.NumberOfRowsToAdd.Value, TempObj.GetFormat().Length]);
                        for (int ri = 0; ri < TempObj.NumberOfRowsToAdd; ri++)
                        {
                            for (int ci = 0; ci < TempObj.GetFormat().Length; ci++)
                            {
                                TempObj.GetContent()[ri, ci] = Content[TotalCount];
                                TotalCount += 1;
                            }
                        }
                        if (TempObj.ContentValidForLoad)
                        {
                            break;
                        }
                        return false;
                    }
                default:
                    return false;
            }
            return true;
        }

        public string HandleTemplateCreateModify(ReadOnlySpan<byte> reportName, ReadOnlySpan<byte> SheetName, string[] header, bool createNew)
        {
            //requires 2 allocations to use for logging.
            string sreportName = Encoding.UTF8.GetString(reportName.ToArray());
            string sSheetName = Encoding.UTF8.GetString(SheetName.ToArray());

            try
            {
                if (createNew)
                {
                    if (TemplateObjectCreateUpdateAppendOverwrite(reportName, SheetName, CommandType.CreateTemplate, header, null))
                    {
                        return ("Successfully Initialized " + sreportName + " with Sheet " + sSheetName);
                    }
                    return ("Failed Initializing Template " + sreportName + " with Sheet " + sSheetName);
                }
                else
                {
                    if (TemplateObjectCreateUpdateAppendOverwrite(reportName, SheetName, CommandType.ModifyTemplate, header, null))
                    {
                        return ("Successfully Updated " + sreportName + " with Sheet " + sSheetName);
                    }                    
                    return ("Failed Updating Template " + sreportName + " with Sheet " + sSheetName);
                }
            }
            catch (ArgumentException ae)
            {
                return (ae.ParamName + " " + ae.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed Template action on " + sreportName + " with Sheet " + sSheetName);
                return ("Failed Initializing Template action on " + sreportName + " with Sheet " + sSheetName);
            }
        }

        public string HandleTemplateAddAppend(ReadOnlySpan<byte> reportName, ReadOnlySpan<byte> SheetName, string[] content, bool upover)
        {

            //requires 2 allocations to use for logging.
            string sreportName = Encoding.UTF8.GetString(reportName.ToArray());
            string sSheetName = Encoding.UTF8.GetString(SheetName.ToArray());

            try
            {
                if (upover)
                {
                    if (TemplateObjectCreateUpdateAppendOverwrite(reportName, SheetName, CommandType.AppendContent, null, content))
                    {
                        return (sreportName + " " + sSheetName + " successfully filled");
                    }
                    return (sreportName + " " + sSheetName + " fill failed");
                }
                else
                {
                    if (TemplateObjectCreateUpdateAppendOverwrite(reportName, SheetName, CommandType.AddContent, null, content))
                    {
                        return (sreportName + " " + sSheetName + " successfully filled");
                    }
                    return (sreportName + " " + sSheetName + " fill failed");
                }
            }
            catch (ArgumentException ax)
            {
                return (sreportName + " fill failed " + ax.ParamName + " " + ax.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, sreportName + " fill failed");
                return (sreportName + " fill failed");
            }
        }

        public void UpdateCreateHeaderOrContent(HttpContextResponder responder, ReadOnlyMemory<byte> sdata)
        {
            bool screate = false;
            byte[] amp = Encoding.UTF8.GetBytes("&");
            if (sdata.Span.IndexOf(Encoding.UTF8.GetBytes("&sheetname=")) > 0)
            {
                ReadOnlySpan<byte> name = sdata.Span.Slice(13);
                name = name.Slice(0, name.IndexOf(amp));

                ReadOnlySpan<byte> sheetName = sdata.Span.Slice(13 + name.Length + 11);
                sheetName = sheetName.Slice(0, sheetName.IndexOf(amp));

                if (sdata.Span.IndexOf(Encoding.UTF8.GetBytes("&header=")) > 0)
                {
                    ReadOnlySpan<byte> header = sdata.Span.Slice(13 + name.Length + 11 + sheetName.Length + 8);
                    if (header.IndexOf(amp) > 0)
                    {
                        header = header.Slice(0, header.IndexOf(amp));
                        screate = sdata.Span.Slice(sdata.Span.LastIndexOf(amp) + 11).SequenceEqual(Encoding.UTF8.GetBytes("true"));
                    }
                    responder.WriteResponse(
                        HandleTemplateCreateModify(
                            name,
                            sheetName,
                            Encoding.UTF8.GetString(header).Split(','),
                            screate));
                }
                else if (sdata.Span.IndexOf(Encoding.UTF8.GetBytes("&content=")) > 0)
                {
                    ReadOnlySpan<byte> content = sdata.Span.Slice(13 + name.Length + 11 + sheetName.Length + 9);
                    if (content.IndexOf(amp) > 0)
                    {
                        content = content.Slice(0, content.IndexOf(amp));
                        screate = sdata.Span.Slice(sdata.Span.LastIndexOf(amp) + 11).SequenceEqual(Encoding.UTF8.GetBytes("true"));
                    }
                    responder.WriteResponse(
                        HandleTemplateAddAppend(
                            name,
                            sheetName,
                            Encoding.UTF8.GetString(content).Split(','),
                            screate));
                }
            }
        }

        public void GetObjectHandlerReports(HttpContextResponder responder, ReadOnlyMemory<byte> Name, bool IsFileBased)
        {
            Stream stream = TemplateToExcelStream(Name.Slice(12), IsFileBased);
            if (stream != null)
            {
                responder.DeliverFile(Name.Span.Slice(12), stream);
            }
            responder.OutputContext.Response.OutputStream.Close();
        }

        public Stream TemplateToExcelStream(ReadOnlyMemory<byte> ReportName, bool IsFileBased)
        {
            Stream stream;
            if (IsFileBased)
            {
                TemplateToFile(ReportName);
                stream = null;
            }
            else
            {
                stream = TemplateToMemory(ReportName);
                if (stream == null)
                {
                    _logger.LogError(new Exception("null stream."), "Report is not setup");
                    return null;
                }
            }
            return stream;
        }

        private void TemplateToFile(ReadOnlyMemory<byte> ReportName)
        {
            if (File.Exists(Path.Combine(Directory.GetCurrentDirectory(), "\\Reports\\" + Encoding.UTF8.GetString(ReportName.ToArray()) + ".xlsx")))
            {
                File.Delete(Path.Combine(Directory.GetCurrentDirectory(), "\\Reports\\" + Encoding.UTF8.GetString(ReportName.ToArray()) + ".xlsx"));
            }
        }

        private Stream TemplateToMemory(ReadOnlyMemory<byte> ReportName)
        {
            MemoryStream memory = null;
            if (Container.GetTemplateSheets(ReportName).Any())
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    foreach (TemplateObject TObj in Container.GetTemplateSheets(ReportName))
                    {
                        if (TObj.DataTable == null)
                        {
                            if (TObj.ProcessTemplate)
                            {
                                //add the worksheet
                                ExcelWorksheet ds = package.Workbook.Worksheets.Add(Encoding.UTF8.GetString(TObj.SheetName.ToArray()));
                                //load the data
                                ds.Cells.LoadFromDataTable(TObj.DataTable, true);
                                //format all filled cells as text
                                ds.Cells[ds.Cells.Start.Address + ":" + ds.Cells.End.Address].Style.Numberformat.Format = "@";

                                ds.Cells.AutoFitColumns(0);
                            }
                            continue;
                        }
                    }
                    memory = new MemoryStream(package.GetAsByteArray(), false);
                }
                return memory;
            }
            return memory;
        }
    }
}