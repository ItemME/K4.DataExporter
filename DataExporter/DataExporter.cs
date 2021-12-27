using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;

namespace DataExporter
{
    public abstract class DataExporter : IDisposable
    {
        protected MemoryStream stream { get; } = new MemoryStream();
        private bool isDone = false;

        public void AddSheet<T>(IEnumerable<T> data, string name = null)
        {
            NewSheet(name);

            var properties = GetProperties(data);

            properties = properties.OrderBy(a => a.Attr.ColumnOrder).ToList();
            // Write headers
            int i = 2;
            foreach (var prop in properties)
            {
                WriteData(prop.Attr.Column, i);
                i++;
            }
            NewLine();

            foreach (var line in data)
            {
                foreach (var prop in properties)
                {
                    WriteData(prop.Prop.GetValue(line), 0);
                }
                NewLine();
            }

            Columns columns = AutoSize(GetSheetData(), properties);
            Worksheet worksheet = new Worksheet();
            worksheet.Append(columns);
            if (!isDone)
                ColorizeSheet(data);
            worksheet.Append(GetSheetData());
            GetWorksheetPart().Worksheet = worksheet;
        }

        private void ColorizeSheet<T>(IEnumerable<T> data)
        {
            WorkbookStylesPart workbookStylesPart1 = GetWorkbookPart().AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac x16r2 xr" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            stylesheet1.AddNamespaceDeclaration("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
            stylesheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");

            var properties = GetProperties(data);

            Fonts fonts1 = new Fonts() { Count = Convert.ToUInt32(properties.Count), KnownFonts = true };
            Fills fills1 = new Fills() { Count = Convert.ToUInt32(properties.Count) };

            // Default values
            fonts1.Append(GenerateFont("Calibri", 11, "000000"));
            fonts1.Append(GenerateFont("Calibri", 11, "000000"));
            fills1.Append(GenerateFill("FFFFFF"));
            fills1.Append(GenerateFill("FFFFFF"));
            fills1.Append(GenerateFill("FFFFFF"));

            foreach (var prop in properties)
            {
                fonts1.Append(GenerateFont(prop.Attr.ColumnFontName, prop.Attr.ColumnFontSize, prop.Attr.ColumnFontColorCode));
                fills1.Append(GenerateFill(prop.Attr.ColumnCellColorCode));
            }

            Borders borders1 = new Borders() { Count = 0U };
            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();
            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);
            borders1.Append(border1);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = Convert.ToUInt32(properties.Count) + 2 };
            CellFormats cellFormats1 = new CellFormats() { Count = Convert.ToUInt32(properties.Count) + 2 };

            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            cellStyleFormats1.Append(cellFormat1);
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            cellFormats1.Append(cellFormat3);

            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = 0U };
            cellStyleFormats1.Append(cellFormat5);
            CellFormat cellFormat44 = new CellFormat() { NumberFormatId = 14U, FormatId = 0U, ApplyNumberFormat = true };
            cellFormats1.Append(cellFormat44);

            int i = 2;
            foreach (var prop in properties)
            {
                CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = Convert.ToUInt32(i), FillId = Convert.ToUInt32(i+1), BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
                cellStyleFormats1.Append(cellFormat2);
                CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = Convert.ToUInt32(i), FillId = Convert.ToUInt32(i+1), BorderId = (UInt32Value)0U, FormatId = Convert.ToUInt32(i) };
                cellFormats1.Append(cellFormat4);

                i++;
            }

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)2U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Columns", FormatId = (UInt32Value)1U, BuiltinId = (UInt32Value)26U };
            CellStyle cellStyle2 = new CellStyle() { Name = "Data", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            cellStyles1.Append(cellStyle2);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium9", DefaultPivotStyle = "PivotStyleLight16" };

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);

            workbookStylesPart1.Stylesheet = stylesheet1;

            isDone = true;
        }

        private DocumentFormat.OpenXml.Spreadsheet.Font GenerateFont(string fontName, int fontSize, string colorCode)
        {
            DocumentFormat.OpenXml.Spreadsheet.Font font1 = new DocumentFormat.OpenXml.Spreadsheet.Font();
            FontSize fontSize1 = new FontSize() { Val = fontSize };
            DocumentFormat.OpenXml.Spreadsheet.Color color1 = new DocumentFormat.OpenXml.Spreadsheet.Color() { Rgb = colorCode };
            FontName fontName1 = new FontName() { Val = fontName };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);

            return font1;
        }

        private Fill GenerateFill(string colorCode)
        {
            Fill fill1 = new Fill();
            PatternFill patternFill1;

            if (colorCode != "NULL")
                patternFill1 = new PatternFill() { PatternType = PatternValues.Solid };
            else
                patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = colorCode };
            patternFill1.Append(foregroundColor1);
            fill1.Append(patternFill1);

            return fill1;
        }

        private Columns AutoSize(SheetData sheetData, List<ColumnInfo> properties)
        {
            var maxColWidth = GetMaxCharacterWidth(sheetData, properties);

            Columns columns = new Columns();
            double maxWidth = 11;
            foreach (var item in maxColWidth)
            {
                double width = Math.Truncate((item.Value * maxWidth + 5) / maxWidth * 256) / 256;
                Column col = new Column() { BestFit = true, Min = (UInt32)(item.Key + 1), Max = (UInt32)(item.Key + 1), CustomWidth = true, Width = (DoubleValue)width };
                columns.Append(col);
            }

            return columns;
        }


        private Dictionary<int, int> GetMaxCharacterWidth(SheetData sheetData, List<ColumnInfo> properties)
        {
            Dictionary<int, int> maxColWidth = new Dictionary<int, int>();
            Dictionary<int, int> maxColWidthTitle = new Dictionary<int, int>();
            Dictionary<int, int> fontSizePerStyle = new Dictionary<int, int>();
            Dictionary<int, string> fontNamePerStyle = new Dictionary<int, string>();

            if (properties == null)
            {
                fontSizePerStyle.Add(0, 11);
                fontNamePerStyle.Add(0, "Calibri");
            }
            else
            {
                int index = 2;
                foreach (var prop in properties)
                {
                    fontSizePerStyle.Add(index, prop.Attr.ColumnFontSize);
                    fontNamePerStyle.Add(index, prop.Attr.ColumnFontName);
                    index++;
                }
            }

            var rows = sheetData.Elements<Row>();
            bool firstRow = true;
            foreach (var r in rows)
            {
                var cells = r.Elements<Cell>().ToArray();

                for (int i = 0; i < cells.Length; i++)
                {
                    var cell = cells[i];
                    var cellValue = cell.CellValue == null ? string.Empty : cell.CellValue.InnerText;
                    var cellTextLength = cellValue.Length;

                    if (firstRow && cell.StyleIndex != null && fontSizePerStyle.ContainsKey(Convert.ToInt32((uint)cell.StyleIndex)))
                    {
                        System.Drawing.Font testFont = new System.Drawing.Font(fontNamePerStyle[Convert.ToInt32((uint)cell.StyleIndex)], fontSizePerStyle[Convert.ToInt32((uint)cell.StyleIndex)]);
                        cellTextLength = (int)Math.Floor(MeasureString(cellValue, testFont).Width) / 6;
                        maxColWidthTitle.Add(i, cellTextLength);
                    }
                    else
                    {
                        if (maxColWidth.ContainsKey(i))
                        {
                            System.Drawing.Font testFont;

                            if (properties != null)
                                testFont = new System.Drawing.Font(fontNamePerStyle[i+2], fontSizePerStyle[i+2]);
                            else
                                testFont = new System.Drawing.Font(fontNamePerStyle[0], fontSizePerStyle[0]);

                            var temp = (int)Math.Floor(MeasureString(cellValue, testFont).Width) / 6;
                            
                            if (temp > maxColWidthTitle[i])
                            {
                                if (temp > maxColWidth[i])
                                    maxColWidth[i] = temp;
                            }
                        }
                        else
                        {
                            System.Drawing.Font testFont;

                            if (properties != null)
                                testFont = new System.Drawing.Font(fontNamePerStyle[i+2], fontSizePerStyle[i+2]);
                            else
                                testFont = new System.Drawing.Font(fontNamePerStyle[0], fontSizePerStyle[0]);

                            var temp = (int)Math.Floor(MeasureString(cellValue, testFont).Width) / 6;

                            if (temp > maxColWidthTitle[i])
                                maxColWidth.Add(i, temp);
                            else
                                maxColWidth.Add(i, maxColWidthTitle[i]);
                        }
                    }
                }

                if (firstRow)
                    firstRow = false;
            }

            return maxColWidth;
        }

        public static SizeF MeasureString(string s, System.Drawing.Font font)
        {
            SizeF result;
            using (var image = new Bitmap(1, 1))
            {
                using (var g = Graphics.FromImage(image))
                {
                    result = g.MeasureString(s, font);
                }
            }

            return result;
        }

        public MemoryStream Export<T>(IEnumerable<T> data)
        {
            AddSheet(data);

            return GetFile();
        }

        public void AddSheet(DataSet reader, string name = null)
        {
            for (int i = 0; i < reader.Tables.Count; i++)
            {
                NewSheet("PAGE " + (i+1));

                foreach (DataColumn column in reader.Tables[i].Columns)
                {
                    WriteData(column.ColumnName, 0);
                }
                NewLine();

                foreach (DataRow row in reader.Tables[i].Rows)
                {
                    foreach (object item in row.ItemArray)
                    {
                        var value = item;
                        if (item == null || item == DBNull.Value)
                        {
                            value = "";
                        }

                        WriteData(value, 0);
                    }
                    NewLine();
                }
  
                Columns columns = AutoSize(GetSheetData(), null);
                Worksheet worksheet = new Worksheet();
                worksheet.Append(columns);
                worksheet.Append(GetSheetData());
                GetWorksheetPart().Worksheet = worksheet;
            }
        }

        public MemoryStream Export(DataSet reader)
        {
            AddSheet(reader);

            return GetFile();
        }

        private List<ColumnInfo> GetProperties<T>(IEnumerable<T> data)
        {
            var typeInfo = typeof(T).GetTypeInfo();

            var properties = new List<ColumnInfo>();
            foreach (var prop in typeInfo.DeclaredProperties)
            {
                var attribute = prop.GetCustomAttribute<ExportColumnAttribute>();
                if (attribute != null)
                {
                    properties.Add(new ColumnInfo(prop, attribute));
                }
            }
            return properties;
        }

        public abstract MemoryStream GetFile();
        protected abstract void NewSheet(string name);
        protected abstract void NewLine();
        protected abstract void WriteData(object data, int styleIndex);
        protected abstract SheetData GetSheetData();
        protected abstract WorksheetPart GetWorksheetPart();
        protected abstract WorkbookPart GetWorkbookPart();

        private class ColumnInfo
        {
            public ColumnInfo(PropertyInfo prop, ExportColumnAttribute attr)
            {
                Prop = prop;
                Attr = attr;
            }

            internal PropertyInfo Prop { get; set; }
            internal ExportColumnAttribute Attr { get; set; }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            stream.Dispose();
        }
    }
}
