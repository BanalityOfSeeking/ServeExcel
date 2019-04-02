using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using Template.Interfaces;

namespace Template.Template
{
    
    public class TemplateObject : ITemplateObject
    {

        public TemplateObject()
        {
        }

        public DataTable DataTable { get; private set; }


        private ReadOnlyMemory<byte> _NameOfReport;
        public ReadOnlyMemory<byte> NameOfReport { get => _NameOfReport; }

        public TemplateObject SetNameOfReport(ReadOnlyMemory<byte> value)
        {
            if (!value.IsEmpty)
            {
                if (value.Length < 4)
                {
                    throw new ArgumentException("Value is not required length", "reportname");
                }
                else
                {
                    _NameOfReport = value;
                    return this;
                }
            }
            throw new ArgumentException("Value is not required length", "reportname");
        }

        private ReadOnlyMemory<byte> _SheetName;
        public ReadOnlyMemory<byte> SheetName { get => _SheetName; }

        public TemplateObject SetSheetName(ReadOnlyMemory<byte> value)
        {
            if (!value.IsEmpty)
            {
                if (value.Length < 5)
                {
                    throw new ArgumentException("Value is not required length of 5 characters", "sheetname");
                }
                else
                {
                    _SheetName = value;
                    return this;
                }
            }
            throw new ArgumentException("Value is not required length of 5 characters", "sheetname");
        }

        private string[] _Format;

        public string[] Format { get => _Format; set => SetFormat(value); }
        public List<string> FormatList { get; set; }

        public int? GetFormatLength()
        {
            return _Format != null ? _Format?.Length : null;
        }

        public string[] GetFormat()
        {
            return _Format != null ? _Format : null;
        }

        public TemplateObject SetFormat(string[] value)
        {
            if (value.Count() > 2)
            {
                _Format = value;
                FormatList = _Format.ToList();                
            }
            else
            {
                throw new ArgumentException("Format has to be greater then 2 elements", "header");
            }
            return this;
        }

        private string[,] _Content { get; set; } = null;

        public string[,] GetContent()
        {
            return _Content;
        }

        public TemplateObject SetContent(string[,] value)
        {
            if (value.GetLength(0) == Format.Length)
            {
                _Content = value;
            }
            else
            {
                throw new ArgumentException("Format must match length of report", "content");
            }
            return this;
        }

        private string[] _ContentArray;

        public string[] GetContentArray()
        {
            return _ContentArray;
        }

        public TemplateObject SetContentArray(string[] value)
        {
            if (value.Length % Format.Length == 0)
            {
                _ContentArray = value;
            }
            else
            {
                throw new ArgumentException("Content array must be multiple of " + Format.Length + ".", nameof(_ContentArray));
            }
            return this;
        }

        public int? ContentArrayLength => ContentArrayAdded ? GetContentArray()?.Length ?? null : null;

        public bool ContentArrayAdded => GetContentArray().Length > 0;

        public int? NumberOfRowsToAdd => ContentArrayAdded ? GetFormatLength().HasValue ? ContentArrayLength / GetFormatLength().Value : null : null;

        public bool ContentValidForLoad => ContentArrayAdded ? GetContentArray().Length != 0 ? GetFormat().Length != 0 ? GetContentArray().Length % GetFormat().Length == 0 ? true : false : false : false : false;

        public bool ProcessTemplate
        {
            get
            {
                if (ContentValidForLoad)
                {
                    DataTable = new DataTable(Encoding.UTF8.GetString(SheetName.Span))
                    {
                        CaseSensitive = false,
                        Locale = CultureInfo.CurrentCulture
                    };
                    DataColumn[] dc = (from li in FormatList
                                                 select new DataColumn(li)).ToArray();

                    DataTable.Columns.AddRange(dc);

                    int Total = 0;

                    for (int i = 0; i < NumberOfRowsToAdd.Value; i++)
                    {
                        object[] oa = new object[GetFormat().Length];
                        for (int x = 0; x < GetFormat().Length; x++)
                        {
                            oa[x] = GetContentArray()[Total];
                            Total += 1;
                        }
                        DataTable.LoadDataRow(oa, false);
                    }
                    return true;
                }
                return false;
            }
        }
    }
}