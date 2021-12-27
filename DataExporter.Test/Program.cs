using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DataExporter.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = Path.GetTempFileName() + Guid.NewGuid().ToString() + ".xlsx";
            var xlsxExporter = new XlsxDataExporter(false);

            var data1 = new List<StatisticsModel>()
            {
                new StatisticsModel{ Checkpoint=5, TextData="Hello", AutoSize=true, Barcode="TEST1", Date=DateTime.Now},
                new StatisticsModel{ Checkpoint=20, TextData="Yoo", AutoSize=false, Barcode="TEST2", Date=DateTime.Now},
                new StatisticsModel{ Checkpoint=10, TextData="This is some text", AutoSize=false, Barcode="TEST3", Date=DateTime.Now},
            };

            var data2 = new List<StatisticsModel>()
            {
                new StatisticsModel{ Checkpoint=30, TextData="T", AutoSize=false, Barcode="TEST4", Date=DateTime.Now},
                new StatisticsModel{ Checkpoint=40, TextData="U", AutoSize=true, Barcode="TEST5", Date=DateTime.UtcNow},
                new StatisticsModel{ Checkpoint=50, TextData="T", AutoSize=true, Barcode="TEST6", Date=DateTime.Now},
            };

            xlsxExporter.AddSheet(data1, "PAGE 1");
            xlsxExporter.AddSheet(data2, "PAGE 2");

            var xlsxResult = xlsxExporter.GetFile();

            FileStream file = new FileStream(path, FileMode.Create, FileAccess.Write);
            xlsxResult.WriteTo(file);
            file.Close();
            xlsxResult.Close();

            ProcessStartInfo ps = new ProcessStartInfo();
            ps.FileName = "excel";
            ps.Arguments = path;
            ps.UseShellExecute = true;
            Process.Start(ps);
        }
    }
}
