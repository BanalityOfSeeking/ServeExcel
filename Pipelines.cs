using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using TemplateToExcelServer.ContextResponder;
using TemplateToExcelServer.TemplateToExcelManager;

namespace PipelineTemplate
{
    public partial class Pipelines
    {
        public Pipelines(ITemplateToExcelManager objectHandler)
        {
            ObjectHandler = objectHandler ?? throw new ArgumentNullException(nameof(objectHandler));
        }

        ITemplateToExcelManager ObjectHandler { get; }

        public ActionBlock<OutputPipelineParams> GetEndPipeline()
        {
            return new(Command =>
                       {
                           ReadOnlyMemory<byte> sdata = Command.Data;
                           string data = Encoding.UTF8.GetString(sdata.ToArray());
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
                                   ObjectHandler.GetObjectHandlerReports(Command.Responder, sdata);
                                   break;
                               case "/?showreports":
                                   StringBuilder sbdata = new StringBuilder();
                                   sbdata.AppendLine("<center><p>Available Reports</p><table>");
                                   foreach (ReadOnlyMemory<byte> report in (from reportKeys in ObjectHandler.Container.ContainerDictionary.Keys
                                                                            select reportKeys).Distinct())
                                   {
                                       sbdata.AppendLine("<tr><td>" + Encoding.UTF8.GetString(report.ToArray()) + "</td></tr>");
                                   }
                                   sbdata.AppendLine("</table></center>");
                                   Command.Responder.WriteResponse(sbdata.ToString());
                                   break;

                               default:
                                   Command.Responder.OutputContext.Response.Close();
                                   break;
                           }
                           Command.Responder.CloseResponse(' '.ToString());
                       });
        }

        public TransformBlock<InitServerFlow, BuildFlow> GetConfigServer()
        {
            return new(Config =>
                            {
                                Config.Listener.Prefixes.Add(Config.URL + ":" + Config.Port + "/");
                                Config.Listener.Start();
                                return new BuildFlow(Config.Listener, 10, new CancellationToken());
                            });
        }

        public TransformBlock<BuildFlow, OutputFlow> GetBuildConnectSet()
        {
            return new(Server =>
                            {
                                var requests = new HashSet<Task>();
                                for (var i = 0; i < Server.MaxConcurrentRequests; i++)
                                {
                                    requests.Add(Server.Listener.GetContextAsync());
                                }
                                return new OutputFlow(Server.Listener, requests, Server.Token);
                            });
        }

        public TransformBlock<OutputFlow, OutputPipelineParams> GetGetOutputParams()
        {
            return new(async ConnectionSet =>
                            {
                                while (!ConnectionSet.Token.IsCancellationRequested)
                                {
                                    var t = await Task.WhenAny(ConnectionSet.Requests);

                                    ConnectionSet.Requests.Add(ConnectionSet.Listener.GetContextAsync());

                                    if (t is Task<HttpListenerContext>)
                                    {
                                        var context = (t as Task<HttpListenerContext>).Result;
                                        ConnectionSet.Requests.Remove(t);
                                        GetEndPipeline().Post(new OutputPipelineParams(new HttpContextResponder(context), Encoding.UTF8.GetBytes(context.Request.RawUrl), ConnectionSet.Token));
                                    }
                                }
                                return default;
                            });
        }
    }
}
