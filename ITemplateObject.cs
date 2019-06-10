using System;
using System.Collections.Generic;
using System.Data;

namespace TemplateToExcelServer.Template
{
    public interface ITemplateObject
    {
        bool ContentArrayAdded { get; }
        int? ContentArrayLength { get; }
        bool ContentValidForLoad { get; }
        DataTable DataTable { get; }
        string[] Format { get; }
        List<string> FormatList { get; }
        ReadOnlyMemory<byte> NameOfReport { get; }
        int? NumberOfRowsToAdd { get; }
        bool ProcessTemplate { get; }
        ReadOnlyMemory<byte> SheetName { get; }

        string[,] GetContent();
        string[] GetContentArray();
        string[] GetFormat();
        int? GetFormatLength();
        TemplateObject SetContent(string[,] Content);
        TemplateObject SetContentArray(string[] value);
        TemplateObject SetFormat(string[] Header);
        TemplateObject SetNameOfReport(ReadOnlyMemory<byte> ReportName);
        TemplateObject SetSheetName(ReadOnlyMemory<byte> SheetName);
    }
}