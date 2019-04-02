using System;
using System.IO;
using System.Net;
using System.Text;
using Template.ContextResponder;

namespace Template.HttpResponder
{
    public class HttpContextResponder : ContextResponder<HttpListenerContext>
    {
        public HttpContextResponder(HttpListenerContext output) : base(output)
        {
            OutputContext = output;
        }

        public override HttpListenerContext OutputContext { get; set; }

        public override bool ContextPresent()
        {
            return OutputContext != null;
        }

        public override void AbortResponse()
        {
            if (ContextPresent())
            {
                OutputContext.Response.Abort();
            }
        }

        public override void WriteResponse(byte[] Message)
        {
            if (ContextPresent())
            {
                OutputContext.Response.OutputStream.Write(Message);
                OutputContext.Response.OutputStream.Flush();
            }
        }

        public override void WriteResponse(ReadOnlySpan<byte> Message)
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

        public override void CloseResponse(byte[] Message)
        {
            if (ContextPresent())
            {
                OutputContext.Response.Close(Message, false);
            }
        }

        public override void CloseResponse(ReadOnlySpan<byte> Message)
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

        public override bool DeliverFile(ReadOnlySpan<byte> reportName, Stream Stream)
        {
            if (Stream == null)
            {
                return false;
            }
            HttpListenerResponse response = OutputContext.Response;
            response.StatusCode = 200;
            response.StatusDescription = "OK";
            response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            response.ContentLength64 = Stream.Length;
            string rn = Encoding.UTF8.GetString(reportName) + ".xlsx";
            response.AddHeader("Content-disposition", "attachment; filename=" + rn);
            response.SendChunked = false;
            Stream.Position = 0;
            Stream.CopyTo(response.OutputStream);
            response.OutputStream.Flush();
            Stream.Close();
            Stream.Dispose();
            return true;
        }
    }
}