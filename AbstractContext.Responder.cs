using System;
using System.IO;
using TemplateToExcelServer.ContextResponder;

namespace TemplateToExcelServer.ContextResponder
{
    public abstract class AbstractContextResponder : IContextResponder
    {
        public abstract void AbortResponse();
        public abstract void CloseResponse(byte[] Message);
        public abstract void CloseResponse(ReadOnlyMemory<byte> Message);
        public abstract void CloseResponse(string Message);
        public abstract bool ContextPresent();
        public abstract bool DeliverFile(ReadOnlyMemory<byte> reportName, Stream Stream);
        public abstract void WriteResponse(byte[] Message);
        public abstract void WriteResponse(ReadOnlyMemory<byte> Message);
        public abstract void WriteResponse(string Message);
    }
}
