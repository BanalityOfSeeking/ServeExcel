using System.Collections.Concurrent;
using System.Collections.Generic;

namespace TemplateToExcelServer.Container
{
    public interface IGenericContainer<T, P>
    {
        ConcurrentDictionary<P, List<(P Name, T Target)>> ContainerDictionary { get; }

        void Add(P MainName, P SubName, T _target);
        IEnumerable<(P SubName, T Target)> GetMain(P MainName);
        T GetTarget(P MainName, P SubName);
        void Remove(P MainName, P SubName);
        void RemoveAll(P MainName);
    }
}