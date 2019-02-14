using System;
using System.Net;
using System.Text;

namespace ServeReports
{
    public class SocketLogger : ISocketLogger
    {
        public ConsoleLogger ServerLog { get; set; }

        public void ClientLog(HttpListenerContext client, string additionalMessage)
        {
            client.Response.OutputStream.Write(Encoding.UTF8.GetBytes(additionalMessage),0, Encoding.UTF8.GetBytes(additionalMessage).Length);
        }
        public void ClientError(HttpListenerContext client, Exception ex, string additionalMessage)
        {
            Console.WriteLine(additionalMessage + " " + ex.Message);
        }
        
    }
}

