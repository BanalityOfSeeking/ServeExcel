using System;

namespace ServeReports
{
    public partial class ServeReports
    {
        [MTAThread]
        private static void Main()
        {
            CreateServer("http://127.0.0.1");
        }
    }
}
