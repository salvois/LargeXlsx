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
    public static class ColumnFormattingTest
    {
        [Test]
        public static void Width()
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                    xlsxWriter.BeginWorksheet("Sheet 1", columns: new[] { XlsxColumn.Formatted(width: 20) });

                using (var package = new ExcelPackage(stream))
                    package.Workbook.Worksheets[0].Column(1).Width.Should().Be(20);
            }
        }

        [Test]
        public static void Style()
        {
            var blueStyle = new XlsxStyle(
                new XlsxFont(XlsxFont.Default.FontName, XlsxFont.Default.FontSize, Color.White),
                new XlsxFill(XlsxFill.Pattern.Solid, Color.FromArgb(0, 0x45, 0x86)),
                XlsxBorder.None,
                XlsxNumberFormat.General);

            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                    xlsxWriter.BeginWorksheet("Sheet 1", columns: new[] { XlsxColumn.Formatted(width: 20, style: blueStyle) });

                using (var package = new ExcelPackage(stream))
                {
                    var style = package.Workbook.Worksheets[0].Column(1).Style;
                    style.Fill.PatternType.Should().Be(ExcelFillStyle.Solid);
                    style.Fill.BackgroundColor.Rgb.Should().Be("004586");
                    style.Font.Color.Rgb.Should().Be("ffffff");
                }
            }
        }

        [Test]
        public static void Hidden()
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                    xlsxWriter.BeginWorksheet("Sheet 1", columns: new[] { XlsxColumn.Formatted(width: 0, hidden: true) });

                using (var package = new ExcelPackage(stream))
                    package.Workbook.Worksheets[0].Column(1).Hidden.Should().BeTrue();
            }
        }

        [Test]
        public static void SkipUnformatted()
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                    xlsxWriter.BeginWorksheet("Sheet 1", columns: new[]
                    {
                        XlsxColumn.Formatted(count: 2, width: 0, hidden: true),
                        XlsxColumn.Unformatted(count: 3),
                        XlsxColumn.Formatted(count: 1, width: 0, hidden: true)
                    });

                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    worksheet.Column(1).Hidden.Should().BeTrue();
                    worksheet.Column(2).Hidden.Should().BeTrue();
                    worksheet.Column(3).Hidden.Should().BeFalse();
                    worksheet.Column(4).Hidden.Should().BeFalse();
                    worksheet.Column(5).Hidden.Should().BeFalse();
                    worksheet.Column(6).Hidden.Should().BeTrue();
                    worksheet.Column(7).Hidden.Should().BeFalse();
                }
            }
        }
    }
}