using System.Collections.Generic;

namespace ServeReports
{
    public class TemplateContainer : ITemplateContainer
    {
        public TemplateContainer()
        {
        }
        private List<TemplateObject> _TObject = new List<TemplateObject>();
        public List<TemplateObject> GetTObject() => _TObject;
        public void AddTObject(TemplateObject templateObject) => _TObject.Add(templateObject);

    }
}

