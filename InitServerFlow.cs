

using System;
using System.Net;

namespace PipelineTemplate
{
    public partial class Pipelines
    {
        public class InitServerFlow : ListenerFlow
        {
            public InitServerFlow(HttpListener listener, string url, int port) : base(listener)
            {
                this.URL = url ?? throw new ArgumentNullException(nameof(url));
                Port = port;
            }
            public string URL { get; }
            public int Port { get; }
        }
    }
}
