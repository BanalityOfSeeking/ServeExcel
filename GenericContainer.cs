using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;

namespace TemplateToExcelServer.Container
{
    public class GenericContainer<T, P> : AbstractContainer<T, P>
    {
        public GenericContainer(ConcurrentDictionary<P, List<(P Name, T Target)>> containerDictionary) : base(containerDictionary)
        {
        }
        public override void Add(P MainName, P SubName, T Target)
        {
            if (ContainerDictionary.ContainsKey(MainName))
            {
                if (!ContainerDictionary[MainName].Exists(x => x.Name.Equals(SubName)))
                {
                    ContainerDictionary[MainName].Add((SubName, Target));
                    return;
                }
            }
        }
        public override IEnumerable<(P SubName, T Target)> GetMain(P MainName)
        {
            if (ContainerDictionary.ContainsKey(MainName))
            {
                return ContainerDictionary[MainName];
            }
            return null;
        }
        public override T GetTarget(P MainName, P SubName)
        {
            if (GetMain(MainName).Where(x => x.SubName.ToString() == SubName.ToString())?.Count() == 1)
            {
                return GetMain(MainName).FirstOrDefault(x => x.SubName.Equals(SubName)).Target;
            }
            return default;

        }
        #nullable enable
        public override void Remove(P? MainName, P? SubName)
        {
            if (MainName is not null && SubName is not null && ContainerDictionary.ContainsKey(MainName))
            {
                ContainerDictionary[MainName].Remove((SubName, GetTarget(MainName, SubName)));
                return;
            }
        }
        public override void RemoveAll(P? MainName)
        {
            if (MainName is not null)
            {
                ContainerDictionary[MainName].Clear();
                return;
            }
        }
    }
    #nullable disable
}
