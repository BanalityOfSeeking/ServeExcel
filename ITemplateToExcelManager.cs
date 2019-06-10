using System;
using System.IO;
using TemplateToExcelServer.Container;
using TemplateToExcelServer.ContextResponder;
using TemplateToExcelServer.Template;

namespace TemplateToExcelServer.TemplateToExcelManager
{
    public interface ITemplateToExcelManager
    {
        IGenericContainer<ITemplateObject, ReadOnlyMemory<byte>> Container { get; }

        void GetObjectHandlerReports(IContextResponder responder, ReadOnlyMemory<byte> Name);
        string HandleTemplateAddAppend(ReadOnlyMemory<byte> reportName, ReadOnlyMemory<byte> SheetName, string[] content, bool upover);
        string HandleTemplateCreateModify(ReadOnlyMemory<byte> reportName, ReadOnlyMemory<byte> SheetName, string[] header, bool createNew);
        bool ObjectCreateUpdateAppendOverwrite(ReadOnlyMemory<byte> ReportName, ReadOnlyMemory<byte> SheetName, TemplateToExcelManager.CommandType CommandType, string[] header, string[] Content);
        Stream TemplateToExcelStream(ReadOnlyMemory<byte> ReportName);
        void UpdateCreateHeaderOrContent(IContextResponder responder, ReadOnlyMemory<byte> sdata);
    }
}