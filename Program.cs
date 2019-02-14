using System;

namespace ServeReports
{
    public partial class ServeReports
    {
        [MTAThread]
        private static void Main()
        {
            ConsoleLogger logger = new ConsoleLogger();
            Server server = new Server("http://127.0.0.1", logger);

        }
    }
}
