using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Template.Interfaces;
using Template.Template;

namespace Template.TemplateContainer
{
    public static class ContainerExtensions
    {
        public static TemplateObject GetTemplateSheet( this TemplateContainer<TemplateObject> container, ReadOnlySpan<byte> ReportName, ReadOnlySpan<byte> SheetName)
        {
            foreach (TemplateObject template in container.GetTemplateSheets(ReportName.ToArray()))
            {
                if (template.SheetName.Span.SequenceEqual(SheetName))
                {
                    return template;
                }
            }
            return null;
        }

        public static IEnumerable<TemplateObject> GetTemplateSheets(this TemplateContainer<TemplateObject> container, ReadOnlyMemory<byte> ReportName)
        {
            string rn = Encoding.UTF8.GetString(ReportName.ToArray());
            return from template in container.GetTemplateReports()
                   where template.NameOfReport.Equals(ReportName)
                   select template;
        }

        public static IEnumerable<TemplateObject> GetTemplateReports(this TemplateContainer<TemplateObject> container)
        {
            return from tObject in container.TemplateObjects
                   select tObject;
        }
    }
    public class TemplateContainer<T> : ITemplateContainer<T>
    {
        private readonly List<T> TempCon = new List<T>();
        public List<T> TemplateObjects => TempCon;

        public void AddTemplate(T templateObject)
        {
            TemplateObjects.Add(templateObject);
        }

        public void RemoveTemplate(T templateObject)
        {
            if (TemplateObjects.Contains(templateObject))
            {
                TemplateObjects.Remove(templateObject);
            }
        }
    }
}