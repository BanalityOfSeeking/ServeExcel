using System;
using System.Collections.Generic;
using System.Data;

namespace Template.Interfaces
{
    public interface IContent
    {
        string[,] GetContent();

        string[] GetContentArray();

        bool ContentArrayAdded { get; }
        int? ContentArrayLength { get; }
        int? NumberOfRowsToAdd { get; }
    }

    public interface ILayout
    {
        string[] GetFormat();

        int? GetFormatLength();

        string[] Format { get; set; }
        List<string> FormatList { get; set; }
    }

    public interface ITemplateObject : ILayout, IContent
    {
        ReadOnlyMemory<byte> NameOfReport { get; }
        ReadOnlyMemory<byte> SheetName { get; }
        DataTable DataTable { get; }

        bool ProcessTemplate { get; }
       
    }
}