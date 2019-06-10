using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Text;
using TemplateToExcelServer.Container;
using TemplateToExcelServer.ContextResponder;
using TemplateToExcelServer.Interfaces;
using TemplateToExcelServer.Template;

namespace TemplateToExcelServer.TemplateToExcelManager
{
    public partial class TemplateToExcelManager : ITemplateToExcelManager
    {
        public readonly ILogger Logger;
        public readonly Encoding Encoder;
        public IGenericContainer<ITemplateObject, ReadOnlyMemory<byte>> Container { get; }
        public TemplateToExcelManager(ILogger logger, IGenericContainer<ITemplateObject, ReadOnlyMemory<byte>> container, Encoding encoding)
        {
            Logger = logger;
            Container = container;
            Encoder = encoding;
        }

        public enum CommandType
        {
            CreateTemplate,
            ModifyTemplate,
            AppendContent,
            AddContent
        }

        public void GetObjectHandlerReports(IContextResponder responder, ReadOnlyMemory<byte> Name)
        {
            var stream = TemplateToExcelStream(Name.Slice(12));
            if (stream != null)
            {
                responder.DeliverFile(Name.Slice(12), stream);
            }
            responder.CloseResponse("");
        }

        public string HandleTemplateAddAppend(ReadOnlyMemory<byte> reportName, ReadOnlyMemory<byte> SheetName, string[] content, bool upover)
        {
            //requires 2 allocations to use for logging.
            var sreportName = Encoder.GetString(reportName.ToArray());
            try
            {
                var sSheetName = Encoder.GetString(SheetName.ToArray());
                if (upover)
                {
                    if (ObjectCreateUpdateAppendOverwrite(reportName, SheetName, CommandType.AppendContent, null, content))
                    {
                        return (sreportName + " " + sSheetName + " successfully filled");
                    }
                    return (sreportName + " " + sSheetName + " fill failed");
                }
                else
                {
                    if (ObjectCreateUpdateAppendOverwrite(reportName, SheetName, CommandType.AddContent, null, content))
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
                Logger.LogError(ex, sreportName + " fill failed");
                return (sreportName + " fill failed");
            }
        }

        public string HandleTemplateCreateModify(ReadOnlyMemory<byte> reportName, ReadOnlyMemory<byte> SheetName, string[] header, bool createNew)
        {
            //requires 2 allocations to use for logging.
            var sreportName = Encoder.GetString(reportName.ToArray());
            var sSheetName = Encoder.GetString(SheetName.ToArray());

            try
            {
                if (createNew)
                {
                    if (ObjectCreateUpdateAppendOverwrite(reportName, SheetName, CommandType.CreateTemplate, header, null))
                    {
                        return ("Successfully Initialized " + sreportName + " with Sheet " + sSheetName);
                    }
                    return ("Failed Initializing Template " + sreportName + " with Sheet " + sSheetName);
                }
                else
                {
                    if (ObjectCreateUpdateAppendOverwrite(reportName, SheetName, CommandType.ModifyTemplate, header, null))
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
                Logger.LogError(ex, "Failed Template action on " + sreportName + " with Sheet " + sSheetName);
                return ("Failed Initializing Template action on " + sreportName + " with Sheet " + sSheetName);
            }
        }

        public bool ObjectCreateUpdateAppendOverwrite(ReadOnlyMemory<byte> ReportName, ReadOnlyMemory<byte> SheetName, CommandType CommandType, string[] header, string[] Content)
        {
            switch (CommandType)
            {
                case CommandType.CreateTemplate://create
                    {
                        if (header == null)
                        {
                            return false;
                        }
                        var TempObj = new TemplateObject()
                            .SetNameOfReport(ReportName.ToArray())
                            .SetSheetName(SheetName.ToArray())
                            .SetFormat(header);

                        Container.Add(ReportName, SheetName, TempObj);
                        break;
                    }
                case CommandType.ModifyTemplate://update
                    {
                        if (Container.GetTarget(ReportName, SheetName)?.SetFormat(header) != null)
                        {
                            break;
                        }
                        return false;
                    }
                case CommandType.AppendContent://append content
                    {
                        var TotalCount = 0;
                        var TempObj = Container.GetTarget(ReportName, SheetName);
                        if (TempObj != null && TempObj.NumberOfRowsToAdd.HasValue)
                        {
                            var lca = TempObj.GetContentArray().ToList();
                            lca.AddRange(Content);
                            TempObj.SetContentArray(lca.ToArray());

                            TempObj.SetContent(new string[TempObj.NumberOfRowsToAdd.Value, TempObj.GetFormat().Length]);
                            for (int ri = TempObj.NumberOfRowsToAdd.Value - 1; ri >= 0; ri--)
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

                        return false;
                    }
                case CommandType.AddContent:
                    {
                        var TotalCount = 0;
                        var TempObj = Container.GetTarget(ReportName, SheetName)?.SetContentArray(Content);
                        TempObj.SetContent(new string[TempObj.NumberOfRowsToAdd.Value, TempObj.GetFormat().Length]);
                        for (int ri = TempObj.NumberOfRowsToAdd.Value - 1; ri >= 0; ri--)
                        {
                            for (int ci = TempObj.GetFormat().Length - 1; ci >= 0; ci--)
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
                    {
                        return false;
                    }
            }
            return true;
        }

        public Stream TemplateToExcelStream(ReadOnlyMemory<byte> ReportName)
        {
            Stream stream;

            stream = TemplateToMemory(ReportName);
            if (stream == null)
            {
                Logger.LogError(new Exception("null stream."), "Report is not setup");
                return null;
            }
            return stream;
        }

        public void UpdateCreateHeaderOrContent(IContextResponder responder, ReadOnlyMemory<byte> sdata)
        {
            var screate = false;
            var amp = Encoder.GetBytes("&");
            if (sdata.Span.IndexOf(Encoder.GetBytes("&sheetname=")) > 0)
            {
                ReadOnlyMemory<byte> name = sdata.Span.Slice(13).Slice(0, sdata.Span.IndexOf(amp)).ToArray().AsMemory();
                ReadOnlyMemory<byte> sheetName = sdata.Span.Slice(13 + name.Length + 11).Slice(0, sdata.Span.IndexOf(amp)).ToArray().AsMemory();

                if (sdata.Span.IndexOf(Encoder.GetBytes("&header=")) > 0)
                {
                    var header = sdata.Span.Slice(13 + name.Length + 11 + sheetName.Length + 8);
                    if (header.IndexOf(amp) > 0)
                    {
                        header = header.Slice(0, header.IndexOf(amp));
                        screate = sdata.Span.Slice(sdata.Span.LastIndexOf(amp) + 11).SequenceEqual(Encoder.GetBytes("true"));
                    }
                    responder.WriteResponse(
                        HandleTemplateCreateModify(
                            name,
                            sheetName,
                            Encoder.GetString(header).Split(','),
                            screate));
                }
                else if (sdata.Span.IndexOf(Encoder.GetBytes("&content=")) > 0)
                {
                    var content = sdata.Span.Slice(13 + name.Length + 11 + sheetName.Length + 9);
                    if (content.IndexOf(amp) > 0)
                    {
                        content = content.Slice(0, content.IndexOf(amp));
                        screate = sdata.Span.Slice(sdata.Span.LastIndexOf(amp) + 11).SequenceEqual(Encoder.GetBytes("true"));
                    }
                    responder.WriteResponse(
                        HandleTemplateAddAppend(
                            name,
                            sheetName,
                            Encoder.GetString(content).Split(','),
                            screate));
                }
            }
        }

        private Stream TemplateToMemory(ReadOnlyMemory<byte> ReportName)
        {
            MemoryStream memory = null;
            if (Container.GetMain(ReportName).Any())
            {
                using (var package = new ExcelPackage())
                {
                    foreach (var TObj in Container.GetMain(ReportName).Select(x => x.Target))
                    {
                        if (TObj.DataTable == null)
                        {
                            if (TObj.ProcessTemplate)
                            {
                                //add the worksheet
                                var ds = package.Workbook.Worksheets.Add(Encoder.GetString(TObj.SheetName.ToArray()));
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