﻿using System;
using System.Data;
using System.IO;
using System.Net;
using System.Text;
using System.Linq;
using System.Threading;


namespace ServeReports
{
    public class Server
    {
        private readonly ILogger _logger;

        private Temphandler Temphandler = new Temphandler(new SocketLogger());

        public Server(string IP, ILogger logger)
        {
            _logger = logger;
            CreateServer(IP);
        }
        private const int PORT = 8183;
        private const int MAX_THREADS = 4;
        private const int DATA_READ_TIMEOUT = 2_000_000;
        private const int STORAGE_SIZE = 1024;

        private struct ThreadParams
        {
            public AutoResetEvent ThreadHandle;
            public HttpListenerContext ClientSocket;
            public int ThreadIndex;
        }

        public void CreateServer(string HTTP_IP)
        {
            WaitHandle[] waitHandles;
            HttpListener listener;

            waitHandles = new WaitHandle[MAX_THREADS];
            for (int i = 0; i < MAX_THREADS; ++i)
            {
                waitHandles[i] = new AutoResetEvent(true);
            }

            listener = new HttpListener();
            listener.Prefixes.Add(HTTP_IP + ":8183/");
            listener.Start();

            while (true)
            {

                HttpListenerContext sock = listener.GetContext();


                int index = WaitHandle.WaitAny(waitHandles);



                ThreadParams context = new ThreadParams()
                {
                    ThreadHandle = (AutoResetEvent)waitHandles[index],
                    ClientSocket = sock,
                    ThreadIndex = index
                };

                ThreadPool.QueueUserWorkItem(ProcessSocketConnection, context);
            }
        }

        private void ProcessSocketConnection(object threadState)
        {
            ThreadParams state = (ThreadParams)threadState;
            byte[] receiveBuffer = Encoding.UTF8.GetBytes(state.ClientSocket.Request.RawUrl);


            DoWork(state.ClientSocket, receiveBuffer);

            Cleanup();
            void Cleanup()
            {
                
                state.ClientSocket.Response.Close();
                receiveBuffer = new byte[STORAGE_SIZE * 10];
          
                state.ThreadHandle.Set();
            }
        }

        private void DoWork(HttpListenerContext client, byte[] data)
        {
            try
            {
                byte[] amp = Encoding.UTF8.GetBytes("&");

                ReadOnlySpan<byte> sdata = data.AsSpan<byte>();

                string StartCommand = Encoding.UTF8.GetString(sdata.Slice(0, sdata.IndexOf(Encoding.UTF8.GetBytes("="))).ToArray());
                switch (StartCommand)

                {
                    case "/?reportname=":
                        if (sdata.IndexOf(Encoding.UTF8.GetBytes("&sheetname=")) > 0)
                        {
                            if (sdata.IndexOf(Encoding.UTF8.GetBytes("&header=")) > 0)
                            {
                                if (sdata.IndexOf(Encoding.UTF8.GetBytes("&createnew=")) > 0)
                                {

                                    ReadOnlySpan<byte> name = sdata.Slice(13);

                                    name = name.Slice(0, name.IndexOf(amp));

                                    ReadOnlySpan<byte> sheetName = sdata.Slice(13 + name.Length + 11);

                                    sheetName = sheetName.Slice(0, sheetName.IndexOf(amp));

                                    ReadOnlySpan<byte> header = sdata.Slice(13 + name.Length + 11 + sheetName.Length + 8);

                                    header = header.Slice(0, header.IndexOf(amp));

                                    bool screate = sdata.Slice(sdata.LastIndexOf(amp) + 11).SequenceEqual(Encoding.UTF8.GetBytes("true")) ? true : false;

                                    Temphandler.TemplateInit(Encoding.UTF8.GetString(name.ToArray()), Encoding.UTF8.GetString(sheetName.ToArray()), Encoding.UTF8.GetString(header.ToArray()).Split(','), screate, client);
                                }
                            }
                            else if (sdata.IndexOf(Encoding.UTF8.GetBytes("&content=")) > 0)
                            {
                                ReadOnlySpan<byte> name = sdata.Slice(13);

                                name = name.Slice(0, name.IndexOf(amp));

                                ReadOnlySpan<byte> sheetName = sdata.Slice(13 + name.Length + 11);

                                sheetName = sheetName.Slice(0, sheetName.IndexOf(amp));

                                ReadOnlySpan<byte> content = sdata.Slice(13 + name.Length + 11 + sheetName.Length + 9);

                                Temphandler.TemplateFill(Encoding.UTF8.GetString(name.ToArray()), Encoding.UTF8.GetString(sheetName.ToArray()), Encoding.UTF8.GetString(content.ToArray()).Split(','), client);
                                Temphandler.AddSheet(Encoding.UTF8.GetString(name.ToArray()), Encoding.UTF8.GetString(sheetName.ToArray()));
                            }
                        }
                        break;

                    case "/?getreport=":
                        {
                            ReadOnlySpan<byte> name = sdata.Slice(12);
                            MemoryStream ms = Temphandler.ToExcel(Encoding.UTF8.GetString(name.ToArray()));
                            long len = ms.Length;
                            Temphandler.DeliverFile(client, Encoding.UTF8.GetString(name.ToArray()) + ".xlsx", ms);
                        }
                        break;
                    case "/?report":
                        {
                            string sbdata = "<center><p>Available Reports</p><table>";
                            foreach(var report in (from reportKeys in Temphandler.InMemContainers
                             select reportKeys.Key.Item1).Distinct())

                            {
                                sbdata += "<tr><td>" + report + "</td></tr>";
                            }
                            sbdata += "</table></center>";
                            client.Response.OutputStream.Write(Encoding.UTF8.GetBytes(sbdata), 0, sbdata.Length);
                        }
                        break;

                    default:
                        break;


                }
            }
            catch (Exception ex)
            {
                while (ex.InnerException != null)
                {
                    ex = ex.InnerException;
                }
                Console.Write(ex.Message);
            }

        }
    }
}