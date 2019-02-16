using System;
using System.Data;
using System.Linq;

namespace ServeReports
{
    [Serializable]
    public class TemplateObject : ITemplateObject
    {
        private string _NameOfReport = "";
        private string _SheetName = "";
        private string[] _Format = new string[0];
        private int? _FormatLength = null;
        private string[,] _Content = new string[0, 0];
        private string[] _ContentArray = new string[0];
        public string NameOfReport { get => _NameOfReport; set => _NameOfReport = value; }
        public string SheetName { get => _SheetName; set => _SheetName = value; }
        public string[] Format { get => _Format; set => _Format = value; }
        public int? FormatLength { get => _Format?.Count() > 0 ? Format?.Length : null; }
        public string[,] Content { get => _Content; set { _Content = value; } }
        public string[] ContentArray { get => _ContentArray; set { _ContentArray = value; } }
        public int? ContentArrayLength { get => ContentArrayAdded ? _ContentArray?.Length > 0 ? _ContentArray?.Length : null : null; }
        public bool ContentArrayAdded => ContentArray.Length > 0;
        public bool ContentAdded => Content.Length > 0;
        public int? NumberOfRowsToAdd => ContentAdded ? FormatLength.HasValue ? ContentArrayLength / FormatLength.Value : null : null; 
        public DataTable DataTable { get; set; }
    }
}

