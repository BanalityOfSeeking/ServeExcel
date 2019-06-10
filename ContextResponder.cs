using System;
using System.IO;
using System.Text;
using TemplateToExcelServer.ContextResponder;

namespace TemplateToExcelServer.ContextResponder
{
    public class ContextResponder : AbstractContextResponder
    {
        public ContextResponder()
        {
        }

        public override void AbortResponse()
        {
            Console.Beep();
        }

        public override void CloseResponse(string Message)
        {
            Console.WriteLine(Message);
        }

        public override void CloseResponse(ReadOnlyMemory<byte> Message)
        {
            Console.WriteLine(Encoding.UTF8.GetString(Message.ToArray()));
        }

        public override void CloseResponse(byte[] Message)
        {
            Console.WriteLine(Encoding.UTF8.GetString(Message));
        }

        public override bool ContextPresent()
        {
            return true;
        }

        public override bool DeliverFile(ReadOnlyMemory<byte> reportName, Stream Stream)
        {
            using (FileStream bw = File.Create(Path.Combine("..\\Files", Encoding.UTF8.GetString(reportName.ToArray()))))
            {
                using (StreamReader sr = new StreamReader(Stream))
                {
                    var bContent = Encoding.UTF8.GetBytes(sr.ReadToEnd());
                    bw.Write(bContent, 0, bContent.Length);
                    bw.Flush();
                }
            }
            return true;
        }

        public override void WriteResponse(string Message)
        {
            Console.WriteLine(Message);
        }

        public override void WriteResponse(ReadOnlyMemory<byte> Message)
        {
            Console.WriteLine(Encoding.UTF8.GetString(Message.ToArray()));
        }

        public override void WriteResponse(byte[] Message)
        {
            Console.WriteLine(Encoding.UTF8.GetString(Message));
        }
    }
}