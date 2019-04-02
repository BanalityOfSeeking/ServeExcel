using System.Collections.Generic;

namespace Template.Interfaces
{
    public interface ITemplateContainer<T>
    {
        List<T> TemplateObjects { get; }

        void AddTemplate(T templateObject);

        void RemoveTemplate(T templateObject);
    }
}