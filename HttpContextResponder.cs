using System;
using System.IO;
using System.Net;
using System.Text;
using TemplateToExcelServer.ContextResponder;

namespace TemplateToExcelServer.ContextResponder
{
    public class HttpContextResponder : AbstractContextResponder
    {
        public HttpContextResponder(HttpListenerContext output)
        {
            OutputContext = output;
        }
        public HttpListenerContext OutputContext { get; }

        public override void AbortResponse()
        {
            if (ContextPresent())
            {
                OutputContext.Response.Abort();
            }
        }

        public override void CloseResponse(byte[] Message)
        {
            if (ContextPresent())
            {
                OutputContext.Response.Close(Message, false);
            }
        }

        public override void CloseResponse(ReadOnlyMemory<byte> Message)
        {
            if (ContextPresent())
            {
                OutputContext.Response.Close(Message.ToArray(), false);
            }
        }

        public override void CloseResponse(string Message)
        {
            if (ContextPresent())
            {
                OutputContext.Response.Close(Encoding.UTF8.GetBytes(Message), false);
            }
        }

        public override bool ContextPresent()
        {
            return OutputContext != null;
        }

        public override bool DeliverFile(ReadOnlyMemory<byte> reportName, Stream Stream)
        {
            if (Stream == null)
            {
                return false;
            }
            var response = OutputContext.Response;
            response.StatusCode = 200;
            response.StatusDescription = "OK";
            response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            response.ContentLength64 = Stream.Length;
            var rn = Encoding.UTF8.GetString(reportName.ToArray()) + ".xlsx";
            response.AddHeader("Content-disposition", "attachment; filename=" + rn);
            response.SendChunked = false;
            Stream.Position = 0;
            Stream.CopyTo(response.OutputStream);
            response.OutputStream.Flush();
            Stream.Close();
            Stream.Dispose();
            return true;
        }

        public override void WriteResponse(byte[] Message)
        {
            if (ContextPresent())
            {
                OutputContext.Response.OutputStream.Write(Message);
                OutputContext.Response.OutputStream.Flush();
            }
        }

        public override void WriteResponse(ReadOnlyMemory<byte> Message)
        {
            if (ContextPresent())
            {
                OutputContext.Response.OutputStream.Write(Message.ToArray());
                OutputContext.Response.OutputStream.Flush();
            }
        }

        public override void WriteResponse(string Message)
        {
            if (ContextPresent())
            {
                OutputContext.Response.OutputStream.Write(Encoding.UTF8.GetBytes(Message));
                OutputContext.Response.OutputStream.Flush();
            }
        }
    }
}