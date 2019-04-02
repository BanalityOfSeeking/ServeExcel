using System;
using System.IO;

namespace Template.Interfaces
{
    public interface IContextResponder
    {
        bool ContextPresent();

        bool DeliverFile(ReadOnlySpan<byte> reportName, Stream Stream);

        void CloseResponse(string Message);

        void CloseResponse(ReadOnlySpan<byte> Message);

        void CloseResponse(byte[] Message);

        void AbortResponse();
    }
}