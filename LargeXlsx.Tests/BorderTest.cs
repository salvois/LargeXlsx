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

namespace LargeXlsx.Tests;

[TestFixture]
public static class BorderTest
{
    [TestCase(XlsxBorder.Style.None, ExcelBorderStyle.None)]
    [TestCase(XlsxBorder.Style.Thin, ExcelBorderStyle.Thin)]
    [TestCase(XlsxBorder.Style.Medium, ExcelBorderStyle.Medium)]
    [TestCase(XlsxBorder.Style.Dashed, ExcelBorderStyle.Dashed)]
    [TestCase(XlsxBorder.Style.Dotted, ExcelBorderStyle.Dotted)]
    [TestCase(XlsxBorder.Style.Thick, ExcelBorderStyle.Thick)]
    [TestCase(XlsxBorder.Style.Double, ExcelBorderStyle.Double)]
    [TestCase(XlsxBorder.Style.Hair, ExcelBorderStyle.Hair)]
    [TestCase(XlsxBorder.Style.MediumDashed, ExcelBorderStyle.MediumDashed)]
    [TestCase(XlsxBorder.Style.DashDot, ExcelBorderStyle.DashDot)]
    [TestCase(XlsxBorder.Style.MediumDashDot, ExcelBorderStyle.MediumDashDot)]
    [TestCase(XlsxBorder.Style.DashDotDot, ExcelBorderStyle.DashDotDot)]
    [TestCase(XlsxBorder.Style.MediumDashDotDot, ExcelBorderStyle.MediumDashDotDot)]
    public static void HorizontalAlignment(XlsxBorder.Style borderStyle, ExcelBorderStyle expected)
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1").BeginRow()
                .Write("Test", XlsxStyle.Default.With(new XlsxBorder(top: new XlsxBorder.Line(Color.DeepPink, borderStyle))));
        using var package = new ExcelPackage(stream);
        var border = package.Workbook.Worksheets[0].Cells["A1"].Style.Border;
        border.Top.Color.Rgb.ShouldBe("FFFF1493");
        border.Top.Style.ShouldBe(expected);
    }


    [Test]
    public static void Defaults()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1").BeginRow().Write("Test");
        using var package = new ExcelPackage(stream);
        var style = package.Workbook.Worksheets[0].Cells["A1"].Style;
        style.Border.Top.Color.Rgb.ShouldBeNull();
        style.Border.Top.Style.ShouldBe(ExcelBorderStyle.None);
        style.Border.Right.Color.Rgb.ShouldBeNull();
        style.Border.Right.Style.ShouldBe(ExcelBorderStyle.None);
        style.Border.Bottom.Color.Rgb.ShouldBeNull();
        style.Border.Bottom.Style.ShouldBe(ExcelBorderStyle.None);
        style.Border.Left.Color.Rgb.ShouldBeNull();
        style.Border.Left.Style.ShouldBe(ExcelBorderStyle.None);
        style.Border.Diagonal.Color.Rgb.ShouldBeNull();
        style.Border.Diagonal.Style.ShouldBe(ExcelBorderStyle.None);
        style.Border.DiagonalDown.ShouldBeFalse();
        style.Border.DiagonalUp.ShouldBeFalse();
    }
}