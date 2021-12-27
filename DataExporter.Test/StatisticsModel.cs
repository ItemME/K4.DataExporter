using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataExporter.Test
{
    public class StatisticsModel
    {
        [ExportColumn(Column = "Number", ColumnOrder = 1, ColumnFontName = "Verdana", ColumnFontColorCode = "FFFFFF", ColumnFontSize = 15, ColumnCellColorCode = "000000")]
        public int Checkpoint { get; set; }

        [ExportColumn(Column = "Text", ColumnOrder = 2, ColumnFontColorCode = "FF0000")]
        public string TextData { get; set; }

        [ExportColumn(Column = "TRUEFALSE_PROP", ColumnOrder = 3, ColumnFontColorCode = "00FF00", ColumnFontSize = 20)]
        public bool AutoSize { get; set; }

        [ExportColumn(Column = "Barcode", ColumnOrder = 4, ColumnFontColorCode = "0000FF", ColumnCellColorCode = "000000")]
        public string Barcode { get; set; }

        [ExportColumn(Column = "DateTime", ColumnOrder = 5, ColumnFontColorCode = "00FFFF", ColumnCellColorCode = "000000")]
        public DateTime Date { get; set; }
    }
}
