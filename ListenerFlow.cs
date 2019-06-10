

using System;
using System.Net;

namespace PipelineTemplate
{
    public partial class Pipelines
    {
        public class ListenerFlow
        {
            public ListenerFlow(HttpListener listener)
            {
                Listener = listener ?? throw new ArgumentNullException(nameof(listener));
            }
            public HttpListener Listener { get; }
        }
    }
}
