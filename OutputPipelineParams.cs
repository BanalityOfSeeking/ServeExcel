
using System;
using System.Threading;
using TemplateToExcelServer.ContextResponder;

namespace PipelineTemplate
{
    public partial class Pipelines
    {

        public class OutputPipelineParams
        {
            public OutputPipelineParams(HttpContextResponder responder, ReadOnlyMemory<byte> data, CancellationToken cancellation)
            {
                Responder = responder ?? throw new ArgumentNullException(nameof(responder));
                Data = data;
                Cancellation = cancellation;
            }

            public HttpContextResponder Responder { get; }
            public ReadOnlyMemory<byte> Data { get; }
            public CancellationToken Cancellation { get; }
        }
    }
}
