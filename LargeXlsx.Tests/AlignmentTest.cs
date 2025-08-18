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
using OfficeOpenXml.Style;
using Shouldly;

namespace LargeXlsx.Tests
{

    [TestFixture]
    public static class AlignmentTest
    {
        [TestCase(XlsxAlignment.Horizontal.General, ExcelHorizontalAlignment.General)]
        [TestCase(XlsxAlignment.Horizontal.Left, ExcelHorizontalAlignment.Left)]
        [TestCase(XlsxAlignment.Horizontal.Center, ExcelHorizontalAlignment.Center)]
        [TestCase(XlsxAlignment.Horizontal.Right, ExcelHorizontalAlignment.Right)]
        [TestCase(XlsxAlignment.Horizontal.Fill, ExcelHorizontalAlignment.Fill)]
        [TestCase(XlsxAlignment.Horizontal.Justify, ExcelHorizontalAlignment.Justify)]
        [TestCase(XlsxAlignment.Horizontal.CenterContinuous, ExcelHorizontalAlignment.CenterContinuous)]
        [TestCase(XlsxAlignment.Horizontal.Distributed, ExcelHorizontalAlignment.Distributed)]
        public static void HorizontalAlignment(XlsxAlignment.Horizontal alignment, ExcelHorizontalAlignment expected)
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                    xlsxWriter.BeginWorksheet("Sheet 1").BeginRow()
                        .Write("Test", XlsxStyle.Default.With(new XlsxAlignment(horizontal: alignment)));
                using (var package = new ExcelPackage(stream))
                    package.Workbook.Worksheets[0].Cells["A1"].Style.HorizontalAlignment.ShouldBe(expected);
            }
        }

        [TestCase(XlsxAlignment.Vertical.Top, ExcelVerticalAlignment.Top)]
        [TestCase(XlsxAlignment.Vertical.Center, ExcelVerticalAlignment.Center)]
        [TestCase(XlsxAlignment.Vertical.Bottom, ExcelVerticalAlignment.Bottom)]
        [TestCase(XlsxAlignment.Vertical.Justify, ExcelVerticalAlignment.Justify)]
        [TestCase(XlsxAlignment.Vertical.Distributed, ExcelVerticalAlignment.Distributed)]
        public static void VerticalAlignment(XlsxAlignment.Vertical alignment, ExcelVerticalAlignment expected)
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                    xlsxWriter.BeginWorksheet("Sheet 1").BeginRow()
                        .Write("Test", XlsxStyle.Default.With(new XlsxAlignment(vertical: alignment)));
                using (var package = new ExcelPackage(stream))
                    package.Workbook.Worksheets[0].Cells["A1"].Style.VerticalAlignment.ShouldBe(expected);
            }
        }

        [Test]
        public static void TextRotation()
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                    xlsxWriter.BeginWorksheet("Sheet 1").BeginRow()
                        .Write("Test", XlsxStyle.Default.With(new XlsxAlignment(textRotation: 45)));
                using (var package = new ExcelPackage(stream))
                    package.Workbook.Worksheets[0].Cells["A1"].Style.TextRotation.ShouldBe(45);
            }
        }

        [Test]
        public static void Indent()
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                    xlsxWriter.BeginWorksheet("Sheet 1").BeginRow()
                        .Write("Test", XlsxStyle.Default.With(new XlsxAlignment(indent: 2)));
                using (var package = new ExcelPackage(stream))
                    package.Workbook.Worksheets[0].Cells["A1"].Style.Indent.ShouldBe(2);
            }
        }

        [Test]
        public static void WrapText()
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                    xlsxWriter.BeginWorksheet("Sheet 1").BeginRow()
                        .Write("Test", XlsxStyle.Default.With(new XlsxAlignment(wrapText: true)));
                using (var package = new ExcelPackage(stream))
                    package.Workbook.Worksheets[0].Cells["A1"].Style.WrapText.ShouldBe(true);
            }
        }

        [Test]
        public static void ShrinkToFit()
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                    xlsxWriter.BeginWorksheet("Sheet 1").BeginRow()
                        .Write("Test", XlsxStyle.Default.With(new XlsxAlignment(shrinkToFit: true)));
                using (var package = new ExcelPackage(stream))
                    package.Workbook.Worksheets[0].Cells["A1"].Style.ShrinkToFit.ShouldBe(true);
            }
        }

        [TestCase(XlsxAlignment.ReadingOrder.ContextDependent, ExcelReadingOrder.ContextDependent)]
        [TestCase(XlsxAlignment.ReadingOrder.LeftToRight, ExcelReadingOrder.LeftToRight)]
        [TestCase(XlsxAlignment.ReadingOrder.RightToLeft, ExcelReadingOrder.RightToLeft)]
        public static void ReadingOrder(XlsxAlignment.ReadingOrder readingOrder, ExcelReadingOrder expected)
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                    xlsxWriter.BeginWorksheet("Sheet 1").BeginRow()
                        .Write("Test", XlsxStyle.Default.With(new XlsxAlignment(readingOrder: readingOrder)));
                using (var package = new ExcelPackage(stream))
                    package.Workbook.Worksheets[0].Cells["A1"].Style.ReadingOrder.ShouldBe(expected);
            }
        }

        [Test]
        public static void Defaults()
        {
            using (var stream = new MemoryStream())
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                {
                    xlsxWriter.BeginWorksheet("Sheet 1").BeginRow().Write("Test");
                }

                stream.Flush();
                stream.Position = 0; // <-- Add this line

                using (var package = new ExcelPackage(stream))
                {
                    var count = package.Workbook.Worksheets.Count;

                    var style = package.Workbook.Worksheets[0].Cells["A1"].Style;
                    style.HorizontalAlignment.ShouldBe(ExcelHorizontalAlignment.General);
                    style.VerticalAlignment.ShouldBe(ExcelVerticalAlignment.Bottom);
                    style.ShrinkToFit.ShouldBeFalse();
                    style.WrapText.ShouldBeFalse();
                    style.TextRotation.ShouldBe(0);
                    style.Indent.ShouldBe(0);
                    style.ReadingOrder.ShouldBe(ExcelReadingOrder.ContextDependent);
                }
            }
        }
    }
}