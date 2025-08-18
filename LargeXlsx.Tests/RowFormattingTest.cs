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
using System.Drawing;
using System.IO;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Shouldly;

namespace LargeXlsx.Tests
{

    [TestFixture]
    public static class RowFormattingTest
    {
        [Test]
        public static void Height()
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                    xlsxWriter.BeginWorksheet("Sheet 1").BeginRow(height: 36.5).Write("Test");
                using (var package = new ExcelPackage(stream))
                    package.Workbook.Worksheets[0].Row(1).Height.ShouldBe(36.5);
            }
        }

        [Test]
        public static void Hidden()
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                    xlsxWriter.BeginWorksheet("Sheet 1").BeginRow(hidden: true).Write("Test");
                using (var package = new ExcelPackage(stream))
                    package.Workbook.Worksheets[0].Row(1).Hidden.ShouldBeTrue();
            }
        }

        [Test]
        public static void Style()
        {
            var blueStyle = new XlsxStyle(
                XlsxFont.Default.With(Color.White),
                new XlsxFill(Color.FromArgb(0, 0x45, 0x86)),
                XlsxBorder.None,
                XlsxNumberFormat.General,
                XlsxAlignment.Default);
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                    xlsxWriter.BeginWorksheet("Sheet 1").BeginRow(style: blueStyle).Write("Test");
                using (var package = new ExcelPackage(stream))
                {
                    var row = package.Workbook.Worksheets[0].Row(1);
                    row.Style.Fill.PatternType.ShouldBe(ExcelFillStyle.Solid);
                    row.Style.Fill.BackgroundColor.Rgb.ShouldBe("FF004586");
                    row.Style.Font.Color.Rgb.ShouldBe("FFFFFFFF");
                }
            }
        }
    }
}