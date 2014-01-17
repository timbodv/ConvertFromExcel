# About
PowerShell cmdlet to convert and Excel file (XLS or XLSX) to a dataset for quick export.

## Pre-requisites
Created with Visual Studio 2013 Express, compiled with .Net 4.5, and using the excellent [ExcelDataReader](https://exceldatareader.codeplex.com/) library.

## Compiling
After extracting the code, download the required nuget packages using a command similar to `Update-Package -Reinstall` and then run `MSBuild.exe /p:Configuration=Release`.

## Acknowledgements
* [ExcelDataReader](https://exceldatareader.codeplex.com/)

# Usage
`````
NAME
ConvertFrom-Excel

SYNTAX
ConvertFrom-Excel [-FileName] <string> [-IncludesHeader]  [<CommonParameters>]

PARAMETERS
-FileName <string>
The path to the Excel file you wish to convert

-IncludesHeader
Whether the Excel spread sheet includes a header row (no by default)

ALIASES
None

REMARKS
None
````