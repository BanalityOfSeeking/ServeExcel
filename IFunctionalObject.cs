using System.Threading.Tasks.Dataflow;

namespace Functional
{
    public interface IFunctionalObject<T> where T : new()
    {
        TransformBlock<T, bool>[] CheckBlocks { get; set; }
        ActionBlock<T>[] ActionBlocks { get; set; }
        TransformBlock<T, T>[] Transform { get; set; }
    }
}