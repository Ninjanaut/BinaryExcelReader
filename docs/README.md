Welcome to the BinaryExcelReader project website! BinaryExcelReader is lightweight C# library to ease loading data from Excel binary (xlsb) file format into DataTable object, based on Microsoft OLE DB Driver. Also supports xls, xlsx and xlsm formats.

> if you don't need to load .XLSB format, consider to use excel reader without OLE DB Driver dependency [Ninjanaut.ExcelReader](https://github.com/Ninjanaut/ExcelReader)

# Installation

from nuget package manager console
```powershell
PM> Install-Package Ninjanaut.BinaryExcelReader
```
from command line
```cmd
> dotnet add package Ninjanaut.BinaryExcelReader
```

| Version | Targets |
|- |- |
| 1.x | .NET 5 |

# Features

* Loading from file path via sheet name.
* Duplicate columns are implicitly allowed.
    * Columns `A, B, B, B1` will be loaded as `A, B, B1, B11` (this is OLE DB Driver default setting).
* Another options might be set via options parameter

| Descriptions                           | Options                   | Defaults  | Notes |
| -                                     | -                         | -         | - |
| Skip top rows                         | HeaderRowIndex            | 0         | Keep in mind that OLE DB driver does not take into account blank rows. For example, if you have 4 additional non-header rows from top and two of them are blank, the header row index is 2. Warning: if the row contains formatting, it is not considered blank.
| Remove empty rows                     | RemoveEmptyRows           | true      | If set to false and the row does not contains anything (even formatting), then the row will not be loaded anyway.
| Limit max columns to load             | MaxColumns                | null      | I recommend setting this value so that you don't accidentally load empty columns. |
| Skip header row                       | HeaderExists              | true      | If set to false, HeaderRowIndex property is ignored.

# Usage

```csharp
using Ninjanaut.IO;

// From file path
var path = @"C:\FooExcel.xlsx";
var sheetName = "Sheet1"
var datatable = BinaryExcelReader.ToDataTable(path, sheetName);
```

you can also use options argument

```csharp
using Ninjanaut.IO;

var path = @"C:\FooExcel.xlsx";
var sheetName = "Sheet1"
var options = new BinaryExcelReaderOptions 
{ 
    // Default settings:
    HeaderExists = true
    HeaderRowIndex = 0,
    RemoveEmptyRows = true,
    MaxColumns = null,
});

var datatable = BinaryExcelReader.ToDataTable(path, sheetName, options);

// The options may be defined within the method.
var datatable = ExcelReader.ToDataTable(path, sheetName, new() { MaxColumns = 5 });
```

# Notes

DataTable object is suitable for this purpose, because you can easily view the read data directly in Visual Studio for debug purposes, create a collection of entities from it or pass datatable as parameter directly into the SQL server stored procedure.

# Contribution

If you would like to contribute to the project, please send a pull request to the dev branch.