using System;
using System.Net;
using System.Text;
using System.Threading.Tasks.Dataflow;
using TemplateToExcelServer.Container;
using TemplateToExcelServer.Interfaces;
using TemplateToExcelServer.Logger;
using TemplateToExcelServer.Template;
using TemplateToExcelServer.TemplateToExcelManager;
using System.Collections.Generic;
using System.Collections.Concurrent;
using static PipelineTemplate.Pipelines;

namespace PipelineTemplate
{
    public partial class PipelineTemplate
    {
        private const int Port = 8183;

        private const string Http_Ip = @"http://127.0.0.1";

        public static void Main()
        {
            ILogger Logger = new Logger();

            Pipelines pipelines = new(new TemplateToExcelManager(
                Logger,
                new GenericContainer<ITemplateObject, ReadOnlyMemory<byte>>(
                    new ConcurrentDictionary<ReadOnlyMemory<byte>, List<(ReadOnlyMemory<byte> Name, ITemplateObject Target)>>()
                    ),
                Encoding.UTF8));
            var ConfigServer = pipelines.GetConfigServer();
            var BuildConnectSet = pipelines.GetBuildConnectSet();
            var GetOutputParams = pipelines.GetGetOutputParams();

            ConfigServer.LinkTo(BuildConnectSet, new DataflowLinkOptions { PropagateCompletion = true });
            BuildConnectSet.LinkTo(GetOutputParams, new DataflowLinkOptions { PropagateCompletion = true });
            ConfigServer.Post(new InitServerFlow(new HttpListener(), Http_Ip, Port));
            ConfigServer.Complete();
            GetOutputParams.Completion.Wait();
           
        }
    }
}
