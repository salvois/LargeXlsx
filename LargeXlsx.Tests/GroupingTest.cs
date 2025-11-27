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
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT of THE USE OF
THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/
using System.IO;
using NUnit.Framework;
using Shouldly;

namespace LargeXlsx.Tests;

[TestFixture]
public class GroupingTest
{
    [Test]
    public void Grouping()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            xlsxWriter
                .BeginWorksheet("Sheet 1")
                .BeginRow().Write("Header 1").Write("Header 2").Write("Header 3")
                .BeginRow(outlineLevel: 1).Write("Test 1").Write("Test 2").Write("Test 3")
                .BeginRow(outlineLevel: 1).Write("Test 4").Write("Test 5").Write("Test 6")
                .BeginRow(collapsed: true, outlineLevel: 1).Write("Test 7").Write("Test 8").Write("Test 9");
        }

        using var package = EPPlusWrapper.Create(stream);
        var sheet = package.Workbook.Worksheets[0];
        sheet.Row(2).OutlineLevel.ShouldBe(1);
        sheet.Row(3).OutlineLevel.ShouldBe(1);
        sheet.Row(4).OutlineLevel.ShouldBe(1);
        sheet.Row(4).Collapsed.ShouldBe(true);
    }
}
