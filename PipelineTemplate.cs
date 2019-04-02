using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using Template;
using Template.HttpResponder;
using Template.Interfaces;
using Template.Logger;
using Template.Template;
using Template.TemplateContainer;
using Template.TemplateHandler;

namespace PipelineTemplate
{
    public class PipelineTemplate
    {
        private const int Port = 8183;

        private const string Http_Ip = "http://127.0.0.1";

        public struct OutputPipelineParams
        {
            public HttpContextResponder Responder { get; set; }
            public ReadOnlyMemory<byte> Data { get; set; }
            public AutoResetEvent WaitHandle { get; set; }
        }

        public static void Main()
        {
            ILogger Logger = new Logger();
            
            

            TemplateObjectHandler ObjectHandler = new TemplateObjectHandler(new TemplateContainer<TemplateObject>(), Logger);

            void EndPipeLine(object Command)
            {
                var inputParams = (OutputPipelineParams)Command;
                EndPipeLine(inputParams);
            }

            void EndPipeline(OutputPipelineParams Command)
            {
                ReadOnlyMemory<byte> sdata = Command.Data;
                string data = null;
                if (sdata.Span.IndexOf(Encoding.UTF8.GetBytes("=")) > 0)
                {
                    data = Encoding.UTF8.GetString(sdata.Slice(0, sdata.Span.IndexOf(Encoding.UTF8.GetBytes("=")) + 1).ToArray());
                }

                switch (data ?? Encoding.UTF8.GetString(sdata.ToArray()))
                {
                    case "/?reportname=":
                        ObjectHandler.UpdateCreateHeaderOrContent(Command.Responder, sdata);
                        break;

                    case "/?getreport=":
                        ObjectHandler.GetObjectHandlerReports(Command.Responder, sdata, false);
                        break;

                    case "/?showreports":
                        StringBuilder sbdata = new StringBuilder();
                        sbdata.AppendLine("<center><p>Available Reports</p><table>");
                        foreach (ReadOnlyMemory<byte> report in (from reportKeys in ObjectHandler.Container.GetTemplateReports()
                                                                 select reportKeys.NameOfReport).Distinct())
                        {
                            sbdata.AppendLine("<tr><td>" + Encoding.UTF8.GetString(report.ToArray()) + "</td></tr>");
                        }
                        sbdata.AppendLine("</table></center>");
                        Command.Responder.OutputContext.Response.Close(Encoding.UTF8.GetBytes(sbdata.ToString()), true);
                        break;

                    default:
                        Command.Responder.OutputContext.Response.Close();
                        break;
                }
                Command.Responder.OutputContext.Response.Close();
                Command.Responder.OutputContext = null;
                Command.Responder = null;
                Command.WaitHandle.Set();
            }

            void GetInput(string HTTP_IP)
            {
                
                WaitHandle[] waitHandles;

                waitHandles = new WaitHandle[4];
                for (int i = 0; i < 4; ++i)
                {
                    waitHandles[i] = new AutoResetEvent(true);
                }              

                OutputPipelineParams output = new OutputPipelineParams();
                HttpListener listener = new HttpListener();

                listener.Prefixes.Add(HTTP_IP + ":" + 8183 + "/");

                listener.Start();

                while (true)
                {
                    HttpListenerContext sock = listener.GetContext();

                    int WaitId = WaitHandle.WaitAny(waitHandles);

                    output.WaitHandle = waitHandles[WaitId] as AutoResetEvent;

                    sock.Response.KeepAlive = false;

                    output.Responder = new HttpContextResponder(sock);

                    output.Data = Encoding.UTF8.GetBytes(sock.Request.RawUrl).AsMemory();

                    EndPipeline(output);
                }
            }
            GetInput(Http_Ip);
        }
    }
}