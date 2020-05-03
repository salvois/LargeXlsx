/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020 Salvatore ISAJA. All rights reserved.

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

using System.Drawing;
using System.IO;
using FluentAssertions;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace LargeXlsx.Tests
{
    [TestFixture]
    public class XlsxWriterTest
    {
        [Test]
        public void InsertionPoint()
        {
            using (var stream = new MemoryStream())
            using (var xlsxWriter = new XlsxWriter(stream))
            {
                xlsxWriter.BeginWorksheet("Sheet1")
                    .BeginRow().Write("A1").Write("B1")
                    .BeginRow().Write("A2");

                xlsxWriter.CurrentRowNumber.Should().Be(2);
                xlsxWriter.CurrentColumnNumber.Should().Be(2);
            }
        }

        [Test]
        public void InsertionPointAfterSkipColumn()
        {
            using (var stream = new MemoryStream())
            using (var xlsxWriter = new XlsxWriter(stream))
            {
                xlsxWriter.BeginWorksheet("Sheet1")
                    .BeginRow().Write("A1").Write("B1")
                    .BeginRow().Write("A2").SkipColumns(2);

                xlsxWriter.CurrentRowNumber.Should().Be(2);
                xlsxWriter.CurrentColumnNumber.Should().Be(4);
            }
        }

        [Test]
        public void InsertionPointAfterSkipRows()
        {
            using (var stream = new MemoryStream())
            using (var xlsxWriter = new XlsxWriter(stream))
            {
                xlsxWriter.BeginWorksheet("Sheet1")
                    .BeginRow().Write("A1").Write("B1")
                    .SkipRows(2);

                xlsxWriter.CurrentRowNumber.Should().Be(3);
                xlsxWriter.CurrentColumnNumber.Should().Be(0);
            }
        }

        [Test]
        public void Simple()
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                {
                    var whiteFont = new XlsxFont("Segoe UI", 9, Color.White, bold: true);
                    var blueFill = new XlsxFill(XlsxFill.Pattern.Solid, Color.FromArgb(0, 0x45, 0x86));
                    var yellowFill = new XlsxFill(XlsxFill.Pattern.Solid, Color.FromArgb(0xff, 0xff, 0x88));
                    var headerStyle = new XlsxStyle(whiteFont, blueFill, XlsxBorder.None, XlsxNumberFormat.General);
                    var highlightStyle = new XlsxStyle(XlsxFont.Default, yellowFill, XlsxBorder.None, XlsxNumberFormat.General);

                    xlsxWriter
                        .BeginWorksheet("Sheet&'<1>\"")
                        .SetDefaultStyle(headerStyle)
                        .BeginRow().Write("Col<1>").Write("Col2").Write("Col&3")
                        .BeginRow().Write().Write("Sub2").Write("Sub3")
                        .SetDefaultStyle(XlsxStyle.Default)
                        .BeginRow().Write("Row3").Write(42).Write(-1, highlightStyle)
                        .BeginRow().Write("Row4").SkipColumns(1).Write(1234)
                        .SkipRows(2)
                        .BeginRow().Write("Row7", columnSpan: 2).Write(3.14159265359);
                }

                using (var package = new ExcelPackage(stream))
                {
                    package.Workbook.Worksheets.Count.Should().Be(1);
                    var sheet = package.Workbook.Worksheets[0];
                    sheet.Name.Should().Be("Sheet&'<1>\"");

                    sheet.Cells["A1"].Value.Should().Be("Col<1>");
                    sheet.Cells["B1"].Value.Should().Be("Col2");
                    sheet.Cells["C1"].Value.Should().Be("Col&3");
                    sheet.Cells["A2"].Value.Should().BeNull();
                    sheet.Cells["B2"].Value.Should().Be("Sub2");
                    sheet.Cells["C2"].Value.Should().Be("Sub3");
                    sheet.Cells["A3"].Value.Should().Be("Row3");
                    sheet.Cells["B3"].Value.Should().Be(42);
                    sheet.Cells["C3"].Value.Should().Be(-1);
                    sheet.Cells["A4"].Value.Should().Be("Row4");
                    sheet.Cells["B4"].Value.Should().BeNull();
                    sheet.Cells["C4"].Value.Should().Be(1234);
                    sheet.Cells["A5"].Value.Should().BeNull();
                    sheet.Cells["A6"].Value.Should().BeNull();
                    sheet.Cells["A7"].Value.Should().Be("Row7");
                    sheet.Cells["B7"].Value.Should().BeNull();
                    sheet.Cells["C7"].Value.Should().Be(3.14159265359);

                    sheet.Cells["A7:B7"].Merge.Should().BeTrue();

                    foreach (var cell in new[] { "A1", "B1", "C1", "A2", "B2", "C2" })
                    {
                        sheet.Cells[cell].Style.Fill.PatternType.Should().Be(ExcelFillStyle.Solid);
                        sheet.Cells[cell].Style.Fill.BackgroundColor.Rgb.Should().Be("004586");
                        sheet.Cells[cell].Style.Font.Bold.Should().BeTrue();
                        sheet.Cells[cell].Style.Font.Color.Rgb.Should().Be("ffffff");
                        sheet.Cells[cell].Style.Font.Name.Should().Be("Segoe UI");
                        sheet.Cells[cell].Style.Font.Size.Should().Be(9);
                    }

                    sheet.Cells["C3"].Style.Fill.PatternType.Should().Be(ExcelFillStyle.Solid);
                    sheet.Cells["C3"].Style.Fill.BackgroundColor.Rgb.Should().Be("ffff88");
                    sheet.Cells["C3"].Style.Font.Bold.Should().BeFalse();
                    sheet.Cells["C3"].Style.Font.Color.Rgb.Should().Be("000000");
                    sheet.Cells["C3"].Style.Font.Name.Should().Be("Calibri");
                    sheet.Cells["C3"].Style.Font.Size.Should().Be(11);

                    foreach (var cell in new[] { "A3", "B3", "A4", "B4", "C4", "A5", "B5", "C5", "A6", "B6", "C6", "A7", "B7", "C7" })
                    {
                        sheet.Cells[cell].Style.Fill.PatternType.Should().Be(ExcelFillStyle.None);
                        sheet.Cells[cell].Style.Font.Bold.Should().BeFalse();
                        sheet.Cells[cell].Style.Font.Color.Rgb.Should().Be("000000");
                        sheet.Cells[cell].Style.Font.Name.Should().Be("Calibri");
                        sheet.Cells[cell].Style.Font.Size.Should().Be(11);
                    }
                }
            }
        }

        [Test]
        public void MultipleSheets()
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                {
                    xlsxWriter
                        .BeginWorksheet("Sheet1")
                        .BeginRow().Write("Sheet1.A1").Write("Sheet1.B1").Write("Sheet1.C1")
                        .BeginRow().AddMergedCell(1, 2).Write("Sheet1.A2").SkipColumns(1).Write("Sheet1.C2")
                        .BeginWorksheet("Sheet2")
                        .BeginRow().AddMergedCell(1, 2).Write("Sheet2.A1").SkipColumns(1).Write("Sheet2.C1")
                        .BeginRow().Write("Sheet2.A2").Write("Sheet2.B2").Write("Sheet2.C2");
                }

                using (var package = new ExcelPackage(stream))
                {
                    package.Workbook.Worksheets.Count.Should().Be(2);

                    var sheet1 = package.Workbook.Worksheets[0];
                    sheet1.Name.Should().Be("Sheet1");
                    sheet1.Cells["A1"].Value.Should().Be("Sheet1.A1");
                    sheet1.Cells["B1"].Value.Should().Be("Sheet1.B1");
                    sheet1.Cells["C1"].Value.Should().Be("Sheet1.C1");
                    sheet1.Cells["A2"].Value.Should().Be("Sheet1.A2");
                    sheet1.Cells["B2"].Value.Should().BeNull();
                    sheet1.Cells["C2"].Value.Should().Be("Sheet1.C2");
                    sheet1.Cells["A2:B2"].Merge.Should().BeTrue();

                    var sheet2 = package.Workbook.Worksheets[1];
                    sheet2.Name.Should().Be("Sheet2");
                    sheet2.Cells["A1"].Value.Should().Be("Sheet2.A1");
                    sheet2.Cells["B1"].Value.Should().BeNull();
                    sheet2.Cells["C1"].Value.Should().Be("Sheet2.C1");
                    sheet2.Cells["A2"].Value.Should().Be("Sheet2.A2");
                    sheet2.Cells["B2"].Value.Should().Be("Sheet2.B2");
                    sheet2.Cells["C2"].Value.Should().Be("Sheet2.C2");
                    sheet2.Cells["A1:B1"].Merge.Should().BeTrue();
                }
            }
        }
    }
}