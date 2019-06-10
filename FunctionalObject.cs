
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks.Dataflow;
namespace Functional
{
    public class FunctionalObject<T> where T : new()
    {
        public static ITargetBlock<N> NullTarget<N>()
        {
            return DataflowBlock.NullTarget<N>();
        }

        public BufferBlock<T> DataBuffer = new BufferBlock<T>(new DataflowBlockOptions { EnsureOrdered = true });

        public BufferBlock<T> IOBuffer = new BufferBlock<T>(new DataflowBlockOptions { EnsureOrdered = true });

        public FunctionalObject(Func<T, T>[] Operations = default, Func<T, T>[] Checks = default)
        {
            if (Operations != default)
            {
                TransformBlocks = new List<TransformBlock<T, T>>();
                Operations.ToList().ForEach(Operation => TransformBlocks.Add(new TransformBlock<T, T>(Operation)));
                IOBuffer.LinkTo(TransformBlocks[0], new DataflowLinkOptions { PropagateCompletion = true, Append = true });
                TransformBlocks[0].LinkTo(DataBuffer, new DataflowLinkOptions { PropagateCompletion = true, Append = true });

                for (var Block = 0; Block < TransformBlocks.Skip(1).Count(); Block ++)
                {
                    DataBuffer.LinkTo(TransformBlocks[Block + 1], new DataflowLinkOptions { PropagateCompletion = true, Append = true });
                    TransformBlocks[Block + 1].LinkTo(IOBuffer, new DataflowLinkOptions { PropagateCompletion = true, Append = true });
                }
            }
            if (Checks != default)
            {
                CheckBlocks = new List<TransformBlock<T, T>>();
                Checks.ToList().ForEach(Check => CheckBlocks.Add(new TransformBlock<T, T>(Check)));
                foreach (var Check in CheckBlocks)
                {
                    IOBuffer.LinkTo(Check, new DataflowLinkOptions { PropagateCompletion = true, Append = true });
                    Check.LinkTo(DataBuffer, new DataflowLinkOptions { PropagateCompletion = true, Append = true });
                    Check.LinkTo(NullTarget<T>(), new DataflowLinkOptions { PropagateCompletion = true, Append = true });
                }
            }
            DataBuffer.LinkTo(NullTarget<T>(), new DataflowLinkOptions { PropagateCompletion = false });
        }

        public List<TransformBlock<T,T>> CheckBlocks { get; }

        public List<TransformBlock<T,T>> TransformBlocks { get; }
    }
}

