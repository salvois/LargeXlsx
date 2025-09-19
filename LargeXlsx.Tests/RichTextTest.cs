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
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Shouldly;

namespace LargeXlsx.Tests;


[TestFixture]
public static class RichTextTest
{
    [Test]
    public static void Write()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            var runs = new XlsxRichTextRun[]
            {
                new("A", XlsxFont.Default.WithSize(16)),
                new(" str", XlsxFont.Default.WithBold()),
                new("ing", XlsxFont.Default.WithItalic()),
            };

            xlsxWriter
                .BeginWorksheet("Sheet1")
                .BeginRow()
                .WriteRichText(runs, XlsxStyle.Default);
        }

        using var package = new ExcelPackage(stream);
        package.Workbook.Worksheets.Count.ShouldBe(1);
        var sheet = package.Workbook.Worksheets[0];

        var cell = sheet.Cells["A1"];
        cell.Value.ShouldBe("A string");
        cell.IsRichText.ShouldBeTrue();
        
        cell.RichText.Count.ShouldBe(3);
        cell.RichText[0].Text.ShouldBe("A");
        cell.RichText[0].Size.ShouldBe(16);
        cell.RichText[0].Bold.ShouldBeFalse();

        cell.RichText[1].Text.ShouldBe(" str");
        cell.RichText[1].Size.ShouldBe((float)XlsxStyle.Default.Font.Size);
        cell.RichText[1].Bold.ShouldBeTrue();

        cell.RichText[2].Text.ShouldBe("ing");
        cell.RichText[2].Size.ShouldBe((float)XlsxStyle.Default.Font.Size);
        cell.RichText[2].Bold.ShouldBeFalse();
        cell.RichText[2].Italic.ShouldBeTrue();
    }
}