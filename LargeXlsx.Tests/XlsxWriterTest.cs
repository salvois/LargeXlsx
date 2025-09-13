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
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Style;
using Shouldly;

namespace LargeXlsx.Tests;

[TestFixture]
public static class XlsxWriterTest
{
    [Test]
    public static void DisposeWriterTwice()
    {
        using var stream = new MemoryStream();
        var xlsxWriter = new XlsxWriter(stream);
        xlsxWriter.BeginWorksheet("Sheet 1").BeginRow().Write("Hello World!");
        xlsxWriter.Dispose();
        Should.NotThrow(() => xlsxWriter.Dispose());
    }

    [Test]
    public static void InsertionPoint()
    {
        using var stream = new MemoryStream();
        using var xlsxWriter = new XlsxWriter(stream);
        xlsxWriter.BeginWorksheet("Sheet1")
            .BeginRow().Write("A1").Write("B1")
            .BeginRow().Write("A2");

        xlsxWriter.CurrentRowNumber.ShouldBe(2);
        xlsxWriter.CurrentColumnNumber.ShouldBe(2);
    }

    [Test]
    public static void InsertionPointAfterSkipColumn()
    {
        using var stream = new MemoryStream();
        using var xlsxWriter = new XlsxWriter(stream);
        xlsxWriter.BeginWorksheet("Sheet1")
            .BeginRow().Write("A1").Write("B1")
            .BeginRow().Write("A2").SkipColumns(2);

        xlsxWriter.CurrentRowNumber.ShouldBe(2);
        xlsxWriter.CurrentColumnNumber.ShouldBe(4);
    }

    [Test]
    public static void InsertionPointAfterSkipRows()
    {
        using var stream = new MemoryStream();
        using var xlsxWriter = new XlsxWriter(stream);
        xlsxWriter.BeginWorksheet("Sheet1")
            .BeginRow().Write("A1").Write("B1")
            .SkipRows(2);

        xlsxWriter.CurrentRowNumber.ShouldBe(3);
        xlsxWriter.CurrentColumnNumber.ShouldBe(0);
    }

    [Test]
    public static void InsertionPointAfterMergedCells()
    {
        using var stream = new MemoryStream();
        using var xlsxWriter = new XlsxWriter(stream);
        xlsxWriter.BeginWorksheet("Sheet1")
            .BeginRow().Write("A1").Write("B1", columnSpan: 3);

        xlsxWriter.CurrentRowNumber.ShouldBe(1);
        xlsxWriter.CurrentColumnNumber.ShouldBe(5);
    }

    [Test]
    public static void Write()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            xlsxWriter
                .BeginWorksheet("Sheet1")
                .BeginRow()
                .Write("A string")
                .Write(123)
                .Write(456.0)
                .Write(789m)
                .Write(new DateTime(2020, 5, 6, 18, 27, 0), XlsxStyle.Default.With(XlsxNumberFormat.ShortDateTime))
                .Write(true)
                .Write();
        }

        using var package = EPPlusWrapper.Create(stream);
        package.Workbook.Worksheets.Count.ShouldBe(1);
        var sheet = package.Workbook.Worksheets[0];
        sheet.Cells["A1"].Value.ShouldBe("A string");
        sheet.Cells["B1"].Value.ShouldBe(123.0);
        sheet.Cells["C1"].Value.ShouldBe(456.0);
        sheet.Cells["D1"].Value.ShouldBe(789.0);
        sheet.Cells["E1"].Value.ShouldBe(new DateTime(2020, 5, 6, 18, 27, 0));
        sheet.Cells["F1"].Value.ShouldBe(true);
        sheet.Cells["G1"].Value.ShouldBeNull();
    }

    [Test]
    public static void Simple()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            var whiteFont = new XlsxFont("Segoe UI", 9, System.Drawing.Color.White, bold: true);
            var blueFill = new XlsxFill(System.Drawing.Color.FromArgb(0, 0x45, 0x86));
            var yellowFill = new XlsxFill(System.Drawing.Color.FromArgb(0xff, 0xff, 0x88));
            var headerStyle = new XlsxStyle(whiteFont, blueFill, XlsxBorder.None, XlsxNumberFormat.General, XlsxAlignment.Default);
            var highlightStyle = XlsxStyle.Default.With(yellowFill);
            var dateStyle = XlsxStyle.Default.With(XlsxNumberFormat.ShortDateTime);

            xlsxWriter
                .BeginWorksheet("Sheet&'<1>\"")
                .SetDefaultStyle(headerStyle)
                .BeginRow().Write("Col<1>").Write("Col2").Write("Col&3")
                .BeginRow().Write().Write("Sub2").Write("Sub3")
                .SetDefaultStyle(XlsxStyle.Default)
                .BeginRow().Write("Row3").Write(42).Write(-1, highlightStyle)
                .BeginRow().Write("Row4").SkipColumns(1).Write(new DateTime(2020, 5, 6, 18, 27, 0), dateStyle)
                .SkipRows(2)
                .BeginRow().Write("Row7", columnSpan: 2).Write(3.14159265359)
                .BeginRow().Write("Row8").Write(false).Write(true);
        }

        using var package = EPPlusWrapper.Create(stream);
        package.Workbook.Worksheets.Count.ShouldBe(1);
        var sheet = package.Workbook.Worksheets[0];
        sheet.Name.ShouldBe("Sheet&'<1>\"");

        sheet.Cells["A1"].Value.ShouldBe("Col<1>");
        sheet.Cells["B1"].Value.ShouldBe("Col2");
        sheet.Cells["C1"].Value.ShouldBe("Col&3");
        sheet.Cells["A2"].Value.ShouldBeNull();
        sheet.Cells["B2"].Value.ShouldBe("Sub2");
        sheet.Cells["C2"].Value.ShouldBe("Sub3");
        sheet.Cells["A3"].Value.ShouldBe("Row3");
        sheet.Cells["B3"].Value.ShouldBe(42);
        sheet.Cells["C3"].Value.ShouldBe(-1);
        sheet.Cells["A4"].Value.ShouldBe("Row4");
        sheet.Cells["B4"].Value.ShouldBeNull();
        sheet.Cells["C4"].Value.ShouldBe(new DateTime(2020, 5, 6, 18, 27, 0));
        sheet.Cells["C4"].Style.Numberformat.NumFmtID.ShouldBe(22);
        sheet.Cells["A5"].Value.ShouldBeNull();
        sheet.Cells["A6"].Value.ShouldBeNull();
        sheet.Cells["A7"].Value.ShouldBe("Row7");
        sheet.Cells["B7"].Value.ShouldBeNull();
        sheet.Cells["C7"].Value.ShouldBe(3.14159265359);
        sheet.Cells["A8"].Value.ShouldBe("Row8");
        sheet.Cells["B8"].Value.ShouldBe(false);
        sheet.Cells["C8"].Value.ShouldBe(true);

        sheet.Cells["A7:B7"].Merge.ShouldBeTrue();

        foreach (var cell in new[] { "A1", "B1", "C1", "A2", "B2", "C2" })
        {
            sheet.Cells[cell].Style.Fill.PatternType.ShouldBe(ExcelFillStyle.Solid);
            sheet.Cells[cell].Style.Fill.BackgroundColor.Rgb.ShouldBe("FF004586");
            sheet.Cells[cell].Style.Font.Bold.ShouldBeTrue();
            sheet.Cells[cell].Style.Font.Color.Rgb.ShouldBe("FFFFFFFF");
            sheet.Cells[cell].Style.Font.Name.ShouldBe("Segoe UI");
            sheet.Cells[cell].Style.Font.Size.ShouldBe(9);
        }

        sheet.Cells["C3"].Style.Fill.PatternType.ShouldBe(ExcelFillStyle.Solid);
        sheet.Cells["C3"].Style.Fill.BackgroundColor.Rgb.ShouldBe("FFFFFF88");
        sheet.Cells["C3"].Style.Font.Bold.ShouldBeFalse();
        sheet.Cells["C3"].Style.Font.Color.Rgb.ShouldBe("FF000000");
        sheet.Cells["C3"].Style.Font.Name.ShouldBe("Calibri");
        sheet.Cells["C3"].Style.Font.Size.ShouldBe(11);

        foreach (var cell in new[] { "A3", "B3", "A4", "B4", "C4", "A5", "B5", "C5", "A6", "B6", "C6", "A7", "B7", "C7" })
        {
            sheet.Cells[cell].Style.Fill.PatternType.ShouldBe(ExcelFillStyle.None);
            sheet.Cells[cell].Style.Font.Bold.ShouldBeFalse();
            sheet.Cells[cell].Style.Font.Color.Rgb.ShouldBe("FF000000");
            sheet.Cells[cell].Style.Font.Name.ShouldBe("Calibri");
            sheet.Cells[cell].Style.Font.Size.ShouldBe(11);
        }
    }

    [Test]
    public static void UnderLine()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            var singleUnderLineStyle =
                XlsxStyle.Default.With(XlsxFont.Default.WithUnderline());
            var doubleUnderLineStyle =
                XlsxStyle.Default.With(XlsxFont.Default.WithUnderline(XlsxFont.Underline.Double));

            xlsxWriter.BeginWorksheet("Sheet1")
                .BeginRow().Write("Row1")
                .BeginRow().Write("Row2", singleUnderLineStyle)
                .BeginRow().Write("Row3", doubleUnderLineStyle);
        }

        using var package = EPPlusWrapper.Create(stream);
        var sheet = package.Workbook.Worksheets[0];

        sheet.Cells["A1"].Style.Font.UnderLine.ShouldBe(false);
        sheet.Cells["A1"].Style.Font.UnderLineType.ShouldBe(ExcelUnderLineType.None);
        sheet.Cells["A2"].Style.Font.UnderLine.ShouldBe(true);
        sheet.Cells["A2"].Style.Font.UnderLineType.ShouldBe(ExcelUnderLineType.Single);
        sheet.Cells["A3"].Style.Font.UnderLine.ShouldBe(true);
        sheet.Cells["A3"].Style.Font.UnderLineType.ShouldBe(ExcelUnderLineType.Double);
    }

    [Test]
    public static void MultipleSheets()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            xlsxWriter
                .BeginWorksheet("Sheet1")
                .BeginRow().Write("Sheet1.A1").Write("Sheet1.B1").Write("Sheet1.C1")
                .BeginRow().Write("Sheet1.A2", columnSpan: 2).Write("Sheet1.C2")
                .BeginWorksheet("Sheet2")
                .BeginRow().AddMergedCell(1, 2).Write("Sheet2.A1").SkipColumns(1).Write("Sheet2.C1")
                .BeginRow().Write("Sheet2.A2").Write("Sheet2.B2").Write("Sheet2.C2");
        }

        using var package = EPPlusWrapper.Create(stream);
        package.Workbook.Worksheets.Count.ShouldBe(2);

        var sheet1 = package.Workbook.Worksheets[0];
        sheet1.Name.ShouldBe("Sheet1");
        sheet1.Cells["A1"].Value.ShouldBe("Sheet1.A1");
        sheet1.Cells["B1"].Value.ShouldBe("Sheet1.B1");
        sheet1.Cells["C1"].Value.ShouldBe("Sheet1.C1");
        sheet1.Cells["A2"].Value.ShouldBe("Sheet1.A2");
        sheet1.Cells["B2"].Value.ShouldBeNull();
        sheet1.Cells["C2"].Value.ShouldBe("Sheet1.C2");
        sheet1.Cells["A2:B2"].Merge.ShouldBeTrue();

        var sheet2 = package.Workbook.Worksheets[1];
        sheet2.Name.ShouldBe("Sheet2");
        sheet2.Cells["A1"].Value.ShouldBe("Sheet2.A1");
        sheet2.Cells["B1"].Value.ShouldBeNull();
        sheet2.Cells["C1"].Value.ShouldBe("Sheet2.C1");
        sheet2.Cells["A2"].Value.ShouldBe("Sheet2.A2");
        sheet2.Cells["B2"].Value.ShouldBe("Sheet2.B2");
        sheet2.Cells["C2"].Value.ShouldBe("Sheet2.C2");
        sheet2.Cells["A1:B1"].Merge.ShouldBeTrue();
    }

    [Test]
    public static void SplitPanesOnMultipleSheets()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            xlsxWriter
                .BeginWorksheet("Sheet1", splitRow: 1, splitColumn: 2)
                .BeginWorksheet("Sheet2", splitRow: 2, splitColumn: 1)
                .BeginWorksheet("OnlyRows", splitRow: 1)
                .BeginWorksheet("OnlyCols", splitColumn: 1);
        }

        using var package = EPPlusWrapper.Create(stream);
        package.Workbook.Worksheets[0].View.ActiveCell.ShouldBe("C2");
        package.Workbook.Worksheets[1].View.ActiveCell.ShouldBe("B3");
        package.Workbook.Worksheets[2].View.ActiveCell.ShouldBe("A2");
        package.Workbook.Worksheets[3].View.ActiveCell.ShouldBe("B1");
    }

    [Test]
    public static void AutoFilter()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter
                .BeginWorksheet("Sheet 1")
                .BeginRow().Write("A1").Write("B1").Write("C1")
                .BeginRow().Write("A2").Write("B2").Write("C2")
                .BeginRow().Write("A3").Write("B3").Write("C3")
                .BeginRow().Write("A4").Write("B4").Write("C4")
                .SetAutoFilter(1, 1, xlsxWriter.CurrentRowNumber, 3);

        using var package = EPPlusWrapper.Create(stream);
        package.Workbook.Worksheets[0].AutoFilterAddress.Address.ShouldBe("A1:C4");
    }

    [Test]
    public static void WorksheetNameTooLong()
    {
        using var stream = new MemoryStream();
        using var xlsxWriter = new XlsxWriter(stream);
        Should.Throw<ArgumentException>(() => xlsxWriter.BeginWorksheet("A very, very, very, long worksheet name exceeding what Excel can handle"));
    }

    [Test]
    public static void DuplicateWorksheetName()
    {
        using var stream = new MemoryStream();
        using var xlsxWriter = new XlsxWriter(stream);
        xlsxWriter.BeginWorksheet("Sheet1");
        Should.Throw<ArgumentException>(() => xlsxWriter.BeginWorksheet("Sheet1"));
    }

    [Theory]
    public static void RightToLeftWorksheet(bool rightToLeft)
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter
                .BeginWorksheet("Sheet 1", rightToLeft: rightToLeft)
                .BeginRow().Write(@"ما هو ""لوريم إيبسوم"" ؟")
                .BeginRow().Write(@"لوريم إيبسوم(Lorem Ipsum) هو ببساطة نص شكلي (بمعنى أن الغاية هي الشكل وليس المحتوى) ويُستخدم في صناعات المطابع ودور النشر.");

        using var package = EPPlusWrapper.Create(stream);
        package.Workbook.Worksheets[0].View.RightToLeft.ShouldBe(rightToLeft);
    }

    [Theory]
    public static void ShowGridlinesWorksheet(bool showGridLines)
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter
                .BeginWorksheet("Sheet 1", showGridLines: showGridLines)
                .BeginRow().Write("Gridlines are hidden in this sheet.");

        using var package = EPPlusWrapper.Create(stream);
        package.Workbook.Worksheets[0].View.ShowGridLines.ShouldBe(showGridLines);
    }

    [Theory]
    public static void ShowHeadersWorksheet(bool showHeaders)
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter
                .BeginWorksheet("Sheet 1", showHeaders: showHeaders)
                .BeginRow().Write("Row and column headers are hidden in this sheet.");

        using var package = EPPlusWrapper.Create(stream);
        package.Workbook.Worksheets[0].View.ShowHeaders.ShouldBe(showHeaders);
    }

    [TestCase(XlsxWorksheetState.Visible, eWorkSheetHidden.Visible)]
    [TestCase(XlsxWorksheetState.Hidden, eWorkSheetHidden.Hidden)]
    [TestCase(XlsxWorksheetState.VeryHidden, eWorkSheetHidden.VeryHidden)]
    public static void WorksheetVisibility(XlsxWorksheetState state, eWorkSheetHidden expected)
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1", state: state).BeginRow().Write("A1");

        using var package = EPPlusWrapper.Create(stream);
        package.Workbook.Worksheets[0].Hidden.ShouldBe(expected);
    }

    [Test]
    public static void SheetProtection()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            xlsxWriter
                .BeginWorksheet("Sheet1")
                .SetSheetProtection(new XlsxSheetProtection("Lorem ipsum", autoFilter: false))
                .BeginRow().Write("A1");
        }

        using var package = EPPlusWrapper.Create(stream);
        package.Workbook.Worksheets.Count.ShouldBe(1);
        var protection = package.Workbook.Worksheets[0].Protection;
        protection.IsProtected.ShouldBeTrue();
        protection.AllowAutoFilter.ShouldBeTrue();
        protection.AllowDeleteColumns.ShouldBeFalse();
        protection.AllowDeleteRows.ShouldBeFalse();
        protection.AllowEditObject.ShouldBeFalse();
        protection.AllowEditScenarios.ShouldBeFalse();
        protection.AllowFormatCells.ShouldBeFalse();
        protection.AllowFormatColumns.ShouldBeFalse();
        protection.AllowFormatRows.ShouldBeFalse();
        protection.AllowInsertColumns.ShouldBeFalse();
        protection.AllowInsertHyperlinks.ShouldBeFalse();
        protection.AllowInsertRows.ShouldBeFalse();
        protection.AllowPivotTables.ShouldBeFalse();
        protection.AllowSelectLockedCells.ShouldBeTrue();
        protection.AllowSelectUnlockedCells.ShouldBeTrue();
        protection.AllowSort.ShouldBeFalse();
    }

    [Test]
    public static void SheetProtection_PasswordTooShort()
    {
        using var stream = new MemoryStream();
        using var xlsxWriter = new XlsxWriter(stream);
        Should.Throw<ArgumentException>(() => xlsxWriter.BeginWorksheet("Sheet1").SetSheetProtection(new XlsxSheetProtection("")));
    }

    [Test]
    public static void SheetProtection_PasswordTooLong()
    {
        using var stream = new MemoryStream();
        using var xlsxWriter = new XlsxWriter(stream);
        Should.Throw<ArgumentException>(() => xlsxWriter.BeginWorksheet("Sheet1").SetSheetProtection(new XlsxSheetProtection("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in")));
    }

    [Test]
    public static void SharedStrings()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            xlsxWriter
                .BeginWorksheet("Sheet1")
                .BeginRow().Write("Lorem ipsum dolor sit amet")
                .BeginRow().WriteSharedString("Lorem ipsum dolor sit amet")
                .BeginRow().WriteSharedString("Lorem ipsum dolor sit amet")
                .BeginRow().WriteSharedString("consectetur adipiscing elit")
                .BeginRow().WriteSharedString("consectetur adipiscing elit")
                .BeginRow().WriteSharedString("Lorem ipsum dolor sit amet");
        }

        using var package = EPPlusWrapper.Create(stream);
        package.Workbook.Worksheets.Count.ShouldBe(1);
        var sheet = package.Workbook.Worksheets[0];
        sheet.Name.ShouldBe("Sheet1");
        sheet.Cells["A1"].Value.ShouldBe("Lorem ipsum dolor sit amet");
        sheet.Cells["A2"].Value.ShouldBe("Lorem ipsum dolor sit amet");
        sheet.Cells["A3"].Value.ShouldBe("Lorem ipsum dolor sit amet");
        sheet.Cells["A4"].Value.ShouldBe("consectetur adipiscing elit");
        sheet.Cells["A5"].Value.ShouldBe("consectetur adipiscing elit");
        sheet.Cells["A6"].Value.ShouldBe("Lorem ipsum dolor sit amet");
    }

    [Theory]
    public static void RequireCellReferences(bool requireCellReferences)
    {
        using var stream = new MemoryStream();
#if NETCOREAPP2_1_OR_GREATER
        const bool useZip64 = true;
#else
        // ZIP64 does not work with DocumentFormat.OpenXml on .NET Framework
        const bool useZip64 = false;
#endif
        using (var xlsxWriter = new XlsxWriter(new SharpCompressZipWriter(stream, XlsxCompressionLevel.Excel, useZip64: useZip64), requireCellReferences: requireCellReferences))
        {
            xlsxWriter.BeginWorksheet("Sheet1").BeginRow().Write("Lorem").Write("ipsum");
        }

        using var spreadsheetDocument = SpreadsheetDocument.Open(stream, false);
        var sheetId = spreadsheetDocument.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Single(s => s.Name == "Sheet1").Id!.ToString()!;
        var worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart!.GetPartById(sheetId);
        worksheetPart.Worksheet
            .Descendants<Row>().Single()
            .Descendants<Cell>().Any(c => c.CellReference == "B1")
            .ShouldBe(requireCellReferences);
    }

    [Test]
    public static void InvalidXmlCharacters()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream, skipInvalidCharacters: true))
        {
            xlsxWriter
                .BeginWorksheet("Sheet\01")
                .SetHeaderFooter(new XlsxHeaderFooter(
                    oddHeader: "&COdd h\0eader",
                    oddFooter: "&COdd f\0ooter",
                    evenHeader: "&CEven h\0eader",
                    evenFooter: "&CEven f\0ooter",
                    firstHeader: "&CFirst he\0ader",
                    firstFooter: "&CFirst fo\0oter"))
                .BeginRow().Write("Inline str\0ing")
                .BeginRow().WriteSharedString("Shared str\0ing")
                .BeginRow().WriteFormula("=A2&\0A3", result: "Inline stringShared string")
                .BeginRow().AddDataValidation(
                    XlsxDataValidation.List(
                        choices: ["Choice\01", "Choice2"],
                        showErrorMessage: true, errorTitle: "Error ti\0tle", error: "A very inform\0ative error message",
                        showInputMessage: true, promptTitle: "Prompt ti\0tle", prompt: "A very enlig\0htening prompt"));
        }

        using var package = EPPlusWrapper.Create(stream);
        package.Workbook.Worksheets.Count.ShouldBe(1);
        var sheet = package.Workbook.Worksheets[0];
        sheet.Name.ShouldBe("Sheet1");
        sheet.Cells["A1"].Value.ShouldBe("Inline string");
        sheet.Cells["A2"].Value.ShouldBe("Shared string");
        sheet.Cells["A3"].Value.ShouldBe("Inline stringShared string");
        sheet.HeaderFooter.OddHeader.CenteredText.ShouldBe("Odd header");
        sheet.HeaderFooter.OddFooter.CenteredText.ShouldBe("Odd footer");
        sheet.HeaderFooter.EvenHeader.CenteredText.ShouldBe("Even header");
        sheet.HeaderFooter.EvenFooter.CenteredText.ShouldBe("Even footer");
        sheet.HeaderFooter.FirstHeader.CenteredText.ShouldBe("First header");
        sheet.HeaderFooter.FirstFooter.CenteredText.ShouldBe("First footer");
        var dataValidation = (ExcelDataValidationList)package.Workbook.Worksheets[0].DataValidations[0];
        dataValidation.ErrorTitle.ShouldBe("Error title");
        dataValidation.Error.ShouldBe("A very informative error message");
        dataValidation.PromptTitle.ShouldBe("Prompt title");
        dataValidation.Prompt.ShouldBe("A very enlightening prompt");
        dataValidation.Formula.Values.ShouldBe(["Choice1", "Choice2"]);
    }
}