![Nuget](https://img.shields.io/nuget/v/sautinsoft.useoffice) ![Nuget](https://img.shields.io/nuget/dt/sautinsoft.useoffice) 
# .NET SDK to convert between Word, Excel, PowerPoint and PDF formats.

![SautinSoft.useoffice](https://camo.githubusercontent.com/fafb859183d8aa71e0453c5073e5a503da19aa080f26b09f4423ef45e511d342/68747470733a2f2f7777772e73617574696e736f66742e636f6d2f6d656469612f6769746875622f752e706e67)

[SautinSoft.useoffice](https://sautinsoft.com/products/useoffice/) is .NET assembly to convert between Word, Excel, PowerPoint and PDF formats.

DOCX to PDF
XLSX to PDF
PPTX to PDF

Support for all MS Office formats.

## Quick links

+ [Developer Guide](https://sautinsoft.com/products/useoffice/examples/)
+ [API Reference](https://sautinsoft.net/help/convert-rtf-html-doc-docx-xls-xlsx-ppt-pptx-to-pdf-net-library/html/N_SautinSoft.htm)

## Top Features

+ [Convert DOCX file to PDF file.](https://sautinsoft.com/products/useoffice/examples/convert-docx-to-pdf-csharp-vb-net.php)
+ [Convert XLSX file to PDF file.](https://sautinsoft.com/products/useoffice/examples/convert-xlsx-to-pdf-csharp-vb-net.php)
+ [Convert RTF file to PDF file.](https://sautinsoft.com/products/useoffice/examples/convert-rtf-to-pdf-csharp-vb-net.php)
+ [Convert PPTX file to PDF file.](https://sautinsoft.com/products/useoffice/examples/convert-pptx-to-pdf-csharp-vb-net.php)

## System Requirement

* .NET Framework 4.6.1 - 4.8
* .NET Core 2.0 - 3.1, .NET 5, 6, 7, 8
* .NET Standard 2.0
* Windows only

## Getting Started with UseOffice .Net

Are you ready to give UseOffice .NET a try? Simply execute `Install-Package sautinsoft.useoffice` from Package Manager Console in Visual Studio to fetch the NuGet package. If you already have UseOffice .NET and want to upgrade the version, please execute `Update-Package sautinsoft.useoffice` to get the latest version.

## Convert DOCX to PDF

```csharp
SautinSoft.UseOffice u = new SautinSoft.UseOffice();
string inpFile = Path.GetFullPath(@"..\..\example.docx");
string outFile = Path.GetFullPath("Result.pdf");
int ret = u.InitWord();
ret = u.ConvertFile(inpFile, outFile, SautinSoft.UseOffice.eDirection.DOCX_to_PDF);
u.CloseWord();
```
## Convert XLSX to PDF

```csharp
SautinSoft.UseOffice u = new SautinSoft.UseOffice();
string inpFile = Path.GetFullPath(@"..\..\example.xlsx");
string outFile = Path.GetFullPath("Result.pdf");
int ret = u.InitExcel();
ret = u.ConvertFile(inpFile, outFile, SautinSoft.UseOffice.eDirection.XSLX_to_PDF);
u.CloseExcel();
```

## Resources

+ **Website:** [www.sautinsoft.com](https://www.sautinsoft.com)
+ **Product Home:** [UseOffice .Net](https://sautinsoft.com/products/useoffice/)
+ [Download SautinSoft.UseOffice](https://sautinsoft.com/products/useoffice/download.php)
+ [Developer Guide](https://sautinsoft.com/products/useoffice/examples/)
+ [API Reference](https://sautinsoft.net/help/convert-rtf-html-doc-docx-xls-xlsx-ppt-pptx-to-pdf-net-library/html/N_SautinSoft.htm)
+ [Support Team](https://sautinsoft.com/support.php)
+ [License](https://sautinsoft.net/help/convert-rtf-html-doc-docx-xls-xlsx-ppt-pptx-to-pdf-net-library/html/license.htm)
