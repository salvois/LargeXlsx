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
using NUnit.Framework;
using OfficeOpenXml;
using Shouldly;
using System.IO;
using System;


namespace LargeXlsx.Tests
{

    public static class PageBreaksTest
    {
        [Test]
        public static void BeforeInsertionPoint()
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                {
                    xlsxWriter
                        .BeginWorksheet("Sheet1")
                        .BeginRow().Write("A1").Write("B1").AddColumnPageBreak().Write("C1")
                        .BeginRow().Write("A2").AddRowPageBreak().Write("B2").Write("C2")
                        .BeginRow().Write("A3").Write("B3").Write("C3");
                }

                using (var package = new ExcelPackage(stream))
                {
                    var worksheetIndex = package.Compatibility.IsWorksheets1Based ? 1 : 0;
                    var sheet = package.Workbook.Worksheets[worksheetIndex];
                    sheet.Row(1).PageBreak.ShouldBeTrue();
                    sheet.Row(2).PageBreak.ShouldBeFalse();
                    sheet.Row(3).PageBreak.ShouldBeFalse();
                    sheet.Column(1).PageBreak.ShouldBeFalse();
                    sheet.Column(2).PageBreak.ShouldBeTrue();
                    sheet.Column(3).PageBreak.ShouldBeFalse();
                }
            }
        }

        [Test]
        public static void Specific()
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                {
                    xlsxWriter
                        .BeginWorksheet("Sheet1")
                        .AddRowPageBreakBefore(3)
                        .AddColumnPageBreakBefore(2);
                }

                using (var package = new ExcelPackage(stream))
                {
                    var worksheetIndex = package.Compatibility.IsWorksheets1Based ? 1 : 0;
                    var sheet = package.Workbook.Worksheets[worksheetIndex];
                    sheet.Row(1).PageBreak.ShouldBeFalse();
                    sheet.Row(2).PageBreak.ShouldBeTrue();
                    sheet.Row(3).PageBreak.ShouldBeFalse();
                    sheet.Column(1).PageBreak.ShouldBeTrue();
                    sheet.Column(2).PageBreak.ShouldBeFalse();
                    sheet.Column(3).PageBreak.ShouldBeFalse();
                }
            }
        }

        [Test]
        public static void OutOfRange()
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                {
                    xlsxWriter.BeginWorksheet("Sheet1");
                    Should.Throw<ArgumentOutOfRangeException>(() => xlsxWriter.AddColumnPageBreakBefore(1));
                    Should.Throw<ArgumentOutOfRangeException>(() =>
                        xlsxWriter.AddColumnPageBreakBefore(Limits.MaxColumnCount + 1));
                    Should.Throw<ArgumentOutOfRangeException>(() => xlsxWriter.AddRowPageBreakBefore(1));
                    Should.Throw<ArgumentOutOfRangeException>(() =>
                        xlsxWriter.AddRowPageBreakBefore(Limits.MaxRowCount + 1));
                }
            }
        }
    }
}