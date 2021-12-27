using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;

namespace DataExporter
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportColumnAttribute : Attribute
    {
        public string Column { get; set; } = string.Empty;
        public int ColumnOrder { get; set; }
        public string ColumnFontName { get; set; } = "Calibri";
        public int ColumnFontSize { get; set; } = 11;
        public string ColumnFontColorCode { get; set; } = "000000";
        public string ColumnCellColorCode { get; set; } = "NULL";
    }
}
