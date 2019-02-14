using System;
using System.Collections.Generic;

namespace ServeReports
{
    [Serializable]
    struct TemplateObject
    {
        public string NameOfReport { get; set; }
        public string SheetName { get; set; }
        public string[] Format { get; set; }
        public string[,] Content { get; set; }
    }

    [Serializable]
    class TemplateContainer
    {
        public List<TemplateObject> TObject = new List<TemplateObject>();
        public string[] ContentArray { get; set; }
        public int FormatLength { get; set; }
        public int ContentLength { get; set; }
    }
}

