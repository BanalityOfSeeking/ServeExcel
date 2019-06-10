using System;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks.Dataflow;
using TemplateToExcelServer;
using TemplateToExcelServer.Container;
using TemplateToExcelServer.ContextResponder;
using TemplateToExcelServer.Interfaces;
using TemplateToExcelServer.Logger;
using TemplateToExcelServer.Template;
using TemplateToExcelServer.TemplateToExcelManager;
using Functional;
using System.Collections.Generic;
using System.Threading.Tasks;
using TemplateToExcelServer.Container;
using System.Collections.Concurrent;
using static PipelineTemplate.Pipelines;

namespace PipelineTemplate
{
    public partial class PipelineTemplate
    {
        private const int Port = 8183;

        private const string Http_Ip = "http://127.0.0.1";

        public static void Main()
        {
            ILogger Logger = new Logger();

            Pipelines pipelines = new Pipelines(new TemplateToExcelManager(
                Logger,
                new GenericContainer<ITemplateObject, ReadOnlyMemory<byte>>(
                    new ConcurrentDictionary<ReadOnlyMemory<byte>, List<(ReadOnlyMemory<byte> Name, ITemplateObject Target)>>()
                    ),
                Encoding.UTF8),Encoding.UTF8);
            var ConfigServer = pipelines.ConfigServer;
            var BuildConnectSet = pipelines.BuildConnectSet;
            var GetOutputParams = pipelines.GetOutputParams;

            ConfigServer.LinkTo(pipelines.BuildConnectSet, new DataflowLinkOptions { PropagateCompletion = true });
            BuildConnectSet.LinkTo(GetOutputParams, new DataflowLinkOptions { PropagateCompletion = true });
            ConfigServer.Post(new InitServerFlow(new HttpListener(), Http_Ip, Port));
            ConfigServer.Complete();
            GetOutputParams.Completion.Wait();

        }
    }
}
