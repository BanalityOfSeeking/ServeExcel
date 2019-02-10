using System;
using System.Data;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;

namespace ServeReports
{
    public partial class ServeReports
    {


        private const int PORT = 8183;
        private const int MAX_THREADS = 4;
        private const int DATA_READ_TIMEOUT = 2_000_000;
        private const int STORAGE_SIZE = 1024;

        private struct ThreadParams
        {
            public AutoResetEvent ThreadHandle;
            public HttpListenerContext ClientSocket;
            public int ThreadIndex;
            public DataSet reports;
        }

        public static void CreateServer(string HTTP_IP)
        {
            WaitHandle[] waitHandles;
            HttpListener listener;
            Directory.CreateDirectory(Directory.GetCurrentDirectory() + "\\Configs");
            Directory.CreateDirectory(Directory.GetCurrentDirectory() + "\\Worksheets");
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
                //Console.WriteLine("Waiting for a connection");
                HttpListenerContext sock = listener.GetContext();

                // Console.WriteLine("Got a connection");
                //Console.WriteLine("Waiting for idle thread");
                int index = WaitHandle.WaitAny(waitHandles);

                //Console.WriteLine("Starting new thread to process client");

                ThreadParams context = new ThreadParams()
                {
                    ThreadHandle = (AutoResetEvent)waitHandles[index],
                    ClientSocket = sock,
                    ThreadIndex = index,
                    reports = new DataSet("Report")
                };

                ThreadPool.QueueUserWorkItem(ProcessSocketConnection, context);
            }
        }

        private static void ProcessSocketConnection(object threadState)
        {
            ThreadParams state = (ThreadParams)threadState;
            byte[] receiveBuffer = Encoding.UTF8.GetBytes(state.ClientSocket.Request.RawUrl);


            DoWork(state.ClientSocket, receiveBuffer, state.reports);

            Cleanup();
            void Cleanup()
            {
                state.reports.Dispose();
                state.ClientSocket.Response.Close();
                receiveBuffer = new byte[STORAGE_SIZE * 10];
                state.reports = new DataSet("Report");
                state.ThreadHandle.Set();
            }
        }

        private static void DoWork(HttpListenerContext client, byte[] data, DataSet reports)
        {
            try
            {
                byte[] amp = Encoding.UTF8.GetBytes("&");
                //Controller For API
                ReadOnlySpan<byte> sdata = data.AsSpan<byte>();

                //API Build report template Columns
                //<Parameter> reportname: name of report to create or update
                //<Parameter> header: comma delimenated column names.
                //<Parameter> createnew: true/false to create or update the report

                //the below example will create a new Template Object with the format set to values passed to the header parameter
                //http://127.0.0.1:8183/?reportname=DynamicReport&header=field1,field2,field3,field4&createnew=true

                //the next example matches the existing report by name and updates the header to the new values passed
                //http://127.0.0.1:8183/?reportname=DynamicReport&header=column1,column2,column3,column4&createnew=false

                //API Add report data
                //<Parameter> /?reportname=nameOfReport: name of report to add content to
                //<Parameter> &content=: comma delimenated column values

                //example
                //http://127.0.0.1:8183/?reportname=DynamicReport&content=1,2,3,4,a,b,c,d

                //API GET resulting report by name
                //<Parameter> /?getreport=nameOfReport

                //example
                //http://127.0.0.1:8183/?getreport=DynamicReeport

                //API Query available reports
                //<Parameter> "/?reports" : basic request to list available reports

                //example
                //http://127.0.0.1:8183/?reports

                string StartCommand = Encoding.UTF8.GetString(sdata.Slice(0, sdata.IndexOf(Encoding.UTF8.GetBytes("="))).ToArray());
                switch (StartCommand)
                //if (sdata.StartsWith(Encoding.UTF8.GetBytes("/?reportname=")))
                {
                    case "/?reportname=":
                        if (sdata.IndexOf(Encoding.UTF8.GetBytes("&header=")) > 0)
                        {
                            if (sdata.IndexOf(Encoding.UTF8.GetBytes("&createnew=")) > 0)
                            {
                                //string testing = Encoding.UTF8.GetString(sdata.ToArray());
                                ReadOnlySpan<byte> name = sdata.Slice(13);
                                name = name.Slice(0, name.IndexOf(amp));
                                ReadOnlySpan<byte> header = sdata.Slice(13 + name.Length + 8);
                                header = header.Slice(0, header.IndexOf(amp));
                                bool screate = sdata.Slice(sdata.LastIndexOf(amp) + 11).SequenceEqual(Encoding.UTF8.GetBytes("true")) ? true : false;
                                //Must Overload this method
                                ExcelTemplate.TemplateInit(Encoding.UTF8.GetString(name.ToArray()), Encoding.UTF8.GetString(header.ToArray()).Split(','), screate, client);
                            }
                        }
                        if (sdata.IndexOf(Encoding.UTF8.GetBytes("&content=")) > 0)
                        {
                            ReadOnlySpan<byte> name = sdata.Slice(13);
                            name = name.Slice(0, name.IndexOf(amp));
                            ReadOnlySpan<byte> content = sdata.Slice(13 + name.Length + 9);
                            //Must Overload this method
                            ExcelTemplate.TemplateFill(Encoding.UTF8.GetString(name.ToArray()), Encoding.UTF8.GetString(content.ToArray()).Split(','), client);
                            ExcelTemplate.AddSheet(Encoding.UTF8.GetString(name.ToArray()), ref reports);
                        }
                        break;

                    case "/?getreport=":
                        //else if (sdata.StartsWith(Encoding.UTF8.GetBytes("/?getreport=")))
                        {
                            ReadOnlySpan<byte> name = sdata.Slice(12);
                            ExcelTemplate.DataSetToExcel(reports, Encoding.UTF8.GetString(name.ToArray()));
                            long len = new FileInfo(Directory.GetCurrentDirectory() + "\\Worksheets\\" + Encoding.UTF8.GetString(name.ToArray()) + ".xlsx").Length;
                            ExcelTemplate.DeliverFile(client, Encoding.UTF8.GetString(name.ToArray()) + ".xlsx", int.Parse(len.ToString()));
                        }
                        break;
                    case "/?report":
                        //else if (sdata.StartsWith(Encoding.UTF8.GetBytes("/?report")))
                        {
                            string sbdata = "<center><p>Available Reports</p><table>";
                            foreach (string file in Directory.EnumerateFiles(Directory.GetCurrentDirectory() + "\\Worksheets\\", "*.xlsx"))
                            {
                                sbdata += "<tr><td>" + Path.GetFileName(file) + "</td></tr>";
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