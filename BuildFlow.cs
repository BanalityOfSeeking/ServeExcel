

using System.Net;
using System.Threading;

namespace PipelineTemplate
{
    public partial class Pipelines
    {
        public class BuildFlow : ListenerFlow
        {
            public BuildFlow(HttpListener listener, int maxConcurrentRequests, CancellationToken token) : base(listener)
            {
                MaxConcurrentRequests = maxConcurrentRequests;
                Token = token;
            }
            public int MaxConcurrentRequests { get; }
            public CancellationToken Token { get; }

        }
    }
}
