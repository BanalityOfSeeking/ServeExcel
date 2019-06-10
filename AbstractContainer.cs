using System;
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace TemplateToExcelServer.Container
{
    public abstract class AbstractContainer<T, P> : IGenericContainer<T, P>
    {
        protected AbstractContainer(ConcurrentDictionary<P, List<(P Name, T Target)>> containerDictionary)
        {
            ContainerDictionary = containerDictionary ?? throw new ArgumentNullException(nameof(containerDictionary));
        }

        public ConcurrentDictionary<P, List<(P Name, T Target)>> ContainerDictionary { get; }
        public abstract void Add(P MainName, P SubName, T Target);
        public abstract IEnumerable<(P SubName, T Target)> GetMain(P MainName);
        public abstract T GetTarget(P MainName, P SubName);
        public abstract void Remove(P MainName, P SubName);
        public abstract void RemoveAll(P MainName);

    }
}