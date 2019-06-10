

using System.Collections.Generic;
using System.Net;
using System.Threading;
using System.Threading.Tasks;

namespace PipelineTemplate
{
    public partial class Pipelines
    {
        public class OutputFlow : ListenerFlow
        {
            public OutputFlow(HttpListener listener, HashSet<Task> requests, CancellationToken token) : base(listener)
            {
                Requests = requests;
                Token = token;
            }
            public HashSet<Task> Requests { get; }
            public CancellationToken Token { get; }

        }
    }
}
