using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;


namespace TemplateToExcelServer.Template
{
    public class TemplateObject : ITemplateObject
    {

        private string[] ContentArray;

        public TemplateObject()
        {
        }

        public bool ContentArrayAdded => GetContentArray().Length > 0;

        public int? ContentArrayLength => ContentArrayAdded ? GetContentArray()?.Length ?? null : null;

        public bool ContentValidForLoad
        {
            get
            {
                if (ContentArrayLength.HasValue && NumberOfRowsToAdd.HasValue)
                {
                    return true;
                }
                return false;
            }
        }

        public DataTable DataTable { get; private set; }

        public string[] Format { get; private set; }
        public List<string> FormatList { get; private set; }
        public bool IsCompleted { get; private set; }
        public ReadOnlyMemory<byte> NameOfReport { get; private set; }

        public int? NumberOfRowsToAdd => ContentArrayAdded ? GetFormatLength().HasValue ? ContentArrayLength / GetFormatLength().Value : null : null;
        public bool ProcessTemplate
        {
            get
            {
                if (IsCompleted && ContentValidForLoad)
                {
                    DataTable = new DataTable(Encoding.UTF8.GetString(SheetName.Span))
                    {
                        Locale = CultureInfo.CurrentCulture
                    };

                    FormatList.ForEach(x => DataTable.Columns.Add(x));

                    var Total = 0;

                    for (int i = 0; i < NumberOfRowsToAdd.Value; i++)
                    {
                        var oa = new object[GetFormat().Length];
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
        public ReadOnlyMemory<byte> SheetName { get; private set; }

        private string[,] Content { get; set; }

        public string[,] GetContent()
        {
            return Content;
        }

        public string[] GetContentArray()
        {
            return ContentArray;
        }

        public string[] GetFormat()
        {
            return Format ?? null;
        }

        public int? GetFormatLength()
        {
            return Format != null ? Format?.Length : null;
        }

        public TemplateObject SetContent(string[,] Content)
        {
            if (Content.GetLength(0) != Format.Length)
            {
                throw new ArgumentException("Format must match length of report", nameof(Content));
            }
            else
            {
                this.Content = Content;
            }
            return this;
        }

        public TemplateObject SetContentArray(string[] value)
        {
            if (value.Length % Format.Length != 0)
            {
                throw new ArgumentException("Content array must be multiple of " + Format.Length + ".", nameof(ContentArray));
            }
            else
            {
                ContentArray = value;
            }
            return this;
        }

        public TemplateObject SetFormat(string[] Header)
        {
            if (Header.Count() <= 2)
            {
                throw new ArgumentException("Format has to be greater then 2 elements", nameof(Header));
            }
            else
            {
                Format = Header;
                FormatList = Format.ToList();
            }
            return this;
        }

        public TemplateObject SetNameOfReport(ReadOnlyMemory<byte> ReportName)
        {
            if (!ReportName.IsEmpty && ReportName.Length >= 4)
            {
                NameOfReport = ReportName;
                return this;
            }
            throw new ArgumentException("Value is not required length", nameof(ReportName));
        }

        public TemplateObject SetSheetName(ReadOnlyMemory<byte> SheetName)
        {
            if (!SheetName.IsEmpty && SheetName.Length >= 5)
            {
                this.SheetName = SheetName;
                return this;
            }
            throw new ArgumentException("Value is not required length of 5 characters", nameof(SheetName));
        }
    }
}