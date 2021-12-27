# K4.DataExporter
K4.DataExporter is a .NET Standard 2.1 library made to export any IEnumerable&lt;T> to a XLSX file.
This library is based of [Guillaume Lacasa's work](https://github.com/pericia/Pericia.DataExport), updated for .NET Standard 2.1

## Install 

This library is available on Nuget : [![NuGet](https://img.shields.io/nuget/v/K4.DataExporter.svg)](https://www.nuget.org/packages/K4.DataExporter/)

## How to use

The data exported will be a file with column titles on the first line, and all your data on the following lines.
To export your data, you will need to define how it will be exported

### Export using Attributes on your model

Add the attribute `ExportColumn` on the properties you want to export. The attributes contains a few properties, most important one being : `Column` (the name of the column) and `ColumnOrder` (to sort your columns)
```cs
public class SampleData
{
    [ExportColumn(Column= "Number", ColumnOrder = 1)]
       public int IntData { get; set; }

    [ExportColumn(Column= "Text", ColumnOrder = 2)]
       public string TextData { get; set; }
}
```
You also have `ColumnFontName`, `ColumnFontSize`, `ColumnColorCode` and `ColumnCellColorCode`. Remember that these properties only apply to the column title, not data.
```cs
public class SampleData
{
    [ExportColumn(Column= "Number", ColumnOrder = 1, ColumnFontName = "Verdana", ColumnFontColorCode = "FFFFFF", ColumnFontSize = 15, ColumnCellColorCode = "000000")]
       public int IntData { get; set; }

    [ExportColumn(Column = "Text", ColumnOrder = 2, ColumnFontColorCode = "00FF00", ColumnFontSize = 20)]
       public string TextData { get; set; }
}
```
### Create your exporter

```cs
var data = new List<SampleData>()
{
    new SampleData{ IntData=5, TextData="Hello"},
    new SampleData{ IntData=20, TextData="Yoo"},
    new SampleData{ IntData=10, TextData="This is some text"},
};

var xlsxExporter = new XlsxDataExporter(true);
var xlsxResult = xlsxExporter.Export(data);
```
### Create Xlsx file with several sheets

While the csv exporter will only allow you to export one set of data, with the xlsx exporter you can create several sheets with different data on each.
```cs
var xlsxExporter = new XlsxDataExporter(true);
xlsxExporter.AddSheet(data1, name="sheet title 1");
xlsxExporter.AddSheet(data2, name="sheet title 2");
var xlsxResult = xlsxExporter.GetFile();
```

### Export using DataSet/DataTable

If you want to export a query result without binding it to a model, you can use a `DataSet`, it can work with Entity too
```cs
var xlsxExporter = new XlsxDataExporter(false);
DataSet dataSet = new DataSet();

List<Event> list = _context.Events.ToList();
DataTable dt = ToDataTable<Event>(list);

dataSet.Tables.Add(dt);
xlsxExporter.AddSheet(dataSet);

var xlsxResult = xlsxExporter.GetFile();

------------------------------------

private static DataTable ToDataTable<T>(List<T> items)
{
       DataTable dataTable = new DataTable(typeof(T).Name);

       PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
       foreach (PropertyInfo prop in Props)
       {
            dataTable.Columns.Add(prop.Name);
       }
       foreach (T item in items)
       {
             var values = new object[Props.Length];
             for (int i = 0; i < Props.Length; i++)
             {
                  values[i] = Props[i].GetValue(item, null);
             }
             dataTable.Rows.Add(values);
       }
       return dataTable;
}
```

You won't be able to use model's attributes with this export option, so make sure to set 'new XlsxDataExporter(false)'  on initialization when using this export option.

### Result

The exporters will output a `MemoryStream`. You can directly save it to a file, or return it in a `FileResult` in an MVC website.
