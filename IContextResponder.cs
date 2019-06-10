using System;
using System.IO;

namespace TemplateToExcelServer.ContextResponder
{
    public interface IContextResponder
    {
        void AbortResponse();
        void CloseResponse(byte[] Message);
        void CloseResponse(ReadOnlyMemory<byte> Message);
        void CloseResponse(string Message);
        bool ContextPresent();
        bool DeliverFile(ReadOnlyMemory<byte> reportName, Stream Stream);
        void WriteResponse(byte[] Message);
        void WriteResponse(ReadOnlyMemory<byte> Message);
        void WriteResponse(string Message);
    }
}