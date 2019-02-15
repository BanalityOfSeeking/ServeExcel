using System.IO;
using System.Net;
using System.Text;

namespace ServeReports
{
    class HttpContextResponder : IHttpContextResponder
    {
        public HttpContextResponder(HttpListenerContext listenerContext)
        {
            ListenerContext = listenerContext;
        }
        public HttpListenerContext ListenerContext { get; set; }

        public bool ContextPresent()
        {
            return ListenerContext != null;
        }

        public void WriteResponse(string Message)
        {
            if (ContextPresent())
            {
                ListenerContext.Response.OutputStream.Write(Encoding.UTF8.GetBytes(Message), 0, Encoding.UTF8.GetBytes(Message).Length);
            }
        }
        public bool DeliverFile(string reportName, MemoryStream memoryStream, string mimeHeader = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        {
            using (HttpListenerResponse response = ListenerContext.Response)
            {
                response.StatusCode = 200;
                response.StatusDescription = "OK";
                response.ContentType = mimeHeader;
                response.ContentLength64 = memoryStream.Length;
                response.AddHeader("Content-disposition", "attachment; filename=" + reportName);
                response.SendChunked = false;
                using (BinaryWriter bw = new BinaryWriter(response.OutputStream))
                {
                    byte[] bContent = memoryStream.ToArray();
                    bw.Write(bContent, 0, bContent.Length);
                    bw.Flush();
                    bw.Close();
                    return true;
                }
            }
        }
    }
}
