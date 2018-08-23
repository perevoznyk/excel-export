# excel-export
.NET Core 2 library for generating Excel files without using Microsoft Excel.

# How to install
Now available on [NuGet](https://www.nuget.org/packages/excel-export)!

```
PM> Install-Package excel-export
```

Or please download from [Releases](https://github.com/perevoznyk/excel-export/releases).

# Usage demo

```csharp
using Export.XLS;
using System;
using System.Globalization;
using System.IO;

namespace excel_export_demo
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelDocument document = new ExcelDocument();
            document.UserName = "Perevoznyk";
            document.CodePage = CultureInfo.CurrentCulture.TextInfo.ANSICodePage;

            document.ColumnWidth(0, 120);
            document.ColumnWidth(1, 80);

            document[0, 0].Value = "ExcelWriter Demo";
            document[0, 0].Font = new Font("Tahoma", 10, FontStyle.Bold);
            document[0, 0].ForeColor = ExcelColor.DarkRed;
            document[0, 0].Alignment = Alignment.Centered;
            document[0, 0].BackColor = ExcelColor.Silver;

            document.WriteCell(1, 0, "int");
            document.WriteCell(1, 1, 10);

            document.Cell(2, 0).Value = "double";
            document.Cell(2, 1).Value = 1.5;

            document.Cell(3, 0).Value = "date";
            document.Cell(3, 1).Value = DateTime.Now;
            document.Cell(3, 1).Format = @"dd/mm/yyyy";

            FileStream stream = new FileStream("demo.xls", FileMode.Create);

            document.Save(stream);
            stream.Close();
        }
    }
}

```
