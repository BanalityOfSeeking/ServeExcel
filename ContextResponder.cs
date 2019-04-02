using System;
using System.IO;
using System.Text;
using Template.Interfaces;

namespace Template.ContextResponder
{
    public abstract class ContextResponder<T> : IContextResponder
    {
        public ContextResponder(T output)
        {
            OutputContext = output;
        }

        public abstract T OutputContext { get; set; }

        public virtual bool ContextPresent()
        {
            return OutputContext != null;
        }

        public virtual bool DeliverFile(ReadOnlySpan<byte> reportName, Stream Stream)
        {
            using (FileStream bw = File.Create(Path.Combine("..\\Files", Encoding.UTF8.GetString(reportName))))
            {
                using (StreamReader sr = new StreamReader(Stream))
                {
                    byte[] bContent = Encoding.UTF8.GetBytes(sr.ReadToEnd());
                    bw.Write(bContent, 0, bContent.Length);
                    bw.Flush();
                }
            }
            return true;
        }

        public virtual void WriteResponse(string Message)
        {
            Console.WriteLine(Message);
        }

        public virtual void WriteResponse(ReadOnlySpan<byte> Message)
        {
            Console.WriteLine(Encoding.UTF8.GetString(Message));
        }

        public virtual void WriteResponse(byte[] Message)
        {
            Console.WriteLine(Encoding.UTF8.GetString(Message));
        }

        public virtual void CloseResponse(string Message)
        {
            Console.WriteLine(Message);
        }

        public virtual void CloseResponse(ReadOnlySpan<byte> Message)
        {
            Console.WriteLine(Encoding.UTF8.GetString(Message));
        }

        public virtual void CloseResponse(byte[] Message)
        {
            Console.WriteLine(Encoding.UTF8.GetString(Message));
        }

        public virtual void AbortResponse()
        {
            Console.Beep();
        }
    }
}