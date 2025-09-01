/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2025 Salvatore ISAJA. All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice,
this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation
and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED THE COPYRIGHT HOLDER ``AS IS'' AND ANY EXPRESS
OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN
NO EVENT SHALL THE COPYRIGHT HOLDER BE LIABLE FOR ANY DIRECT,
INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Shouldly;
using Color = System.Drawing.Color;

namespace LargeXlsx.Tests;

[TestFixture]
public static class ColumnFormattingTest
{
    [Test]
    public static void Width()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1", columns: new[] { XlsxColumn.Formatted(width: 20) });
        using (var package = EPPlusWrapper.Create(stream))
            package.Workbook.Worksheets[0].Column(1).Width.ShouldBe(20);
    }

    [Test]
    public static void Style()
    {
        var blueStyle = new XlsxStyle(XlsxFont.Default.With(Color.White), new XlsxFill(Color.FromArgb(0, 0x45, 0x86)), XlsxBorder.None, XlsxNumberFormat.General, XlsxAlignment.Default);
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1", columns: new[] { XlsxColumn.Formatted(width: 20, style: blueStyle) });
        using var package = EPPlusWrapper.Create(stream);
        var style = package.Workbook.Worksheets[0].Column(1).Style;
        style.Fill.PatternType.ShouldBe(ExcelFillStyle.Solid);
        style.Fill.BackgroundColor.Rgb.ShouldBe("FF004586");
        style.Font.Color.Rgb.ShouldBe("FFFFFFFF");
    }

    [Test]
    public static void Hidden()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1", columns: new[] { XlsxColumn.Formatted(width: 0, hidden: true) });
        using (var package = EPPlusWrapper.Create(stream))
            package.Workbook.Worksheets[0].Column(1).Hidden.ShouldBeTrue();
    }

    [Test]
    public static void SkipUnformatted()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1", columns: new[]
            {
                XlsxColumn.Formatted(count: 2, width: 0, hidden: true),
                XlsxColumn.Unformatted(count: 3),
                XlsxColumn.Formatted(count: 1, width: 0, hidden: true)
            });
        using var package = EPPlusWrapper.Create(stream);
        var worksheet = package.Workbook.Worksheets[0];
        worksheet.Column(1).Hidden.ShouldBeTrue();
        worksheet.Column(2).Hidden.ShouldBeTrue();
        worksheet.Column(3).Hidden.ShouldBeFalse();
        worksheet.Column(4).Hidden.ShouldBeFalse();
        worksheet.Column(5).Hidden.ShouldBeFalse();
        worksheet.Column(6).Hidden.ShouldBeTrue();
        worksheet.Column(7).Hidden.ShouldBeFalse();
    }

    [Test]
    public static void OnlyUnformatted()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1", columns: new[] { XlsxColumn.Unformatted() });

        using var spreadsheetDocument = SpreadsheetDocument.Open(stream, false);
        var sheetId = spreadsheetDocument.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Single(s => s.Name == "Sheet 1").Id!.ToString()!;
        var worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart!.GetPartById(sheetId);
        worksheetPart.Worksheet.Descendants<Columns>().ShouldBeEmpty();
    }
}