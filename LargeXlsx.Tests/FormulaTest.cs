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
using NUnit.Framework;
using OfficeOpenXml;
using Shouldly;

namespace LargeXlsx.Tests;

[TestFixture]
public static class FormulaTest
{
    [Test]
    public static void FormulaWithResult()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1").BeginRow().WriteFormula("41.5+1", result: 42.5);
        using var package = new ExcelPackage(stream);
        // We don't assert package.Workbook.FullCalcOnLoad because EPPlus always sets it to true
        package.Workbook.Worksheets[0].Cells["A1"].Formula.ShouldBe("41.5+1");
        package.Workbook.Worksheets[0].Cells["A1"].Value.ShouldBe("42.5");
    }

    [Test]
    public static void FormulaWithoutResult()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1").BeginRow().WriteFormula("41+1");
        using var package = new ExcelPackage(stream);
        // We don't assert package.Workbook.FullCalcOnLoad because EPPlus always sets it to true
        package.Workbook.Worksheets[0].Cells["A1"].Formula.ShouldBe("41+1");
        package.Workbook.Worksheets[0].Cells["A1"].Value.ShouldBeNull();
    }
}