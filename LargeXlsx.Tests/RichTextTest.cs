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
using System.Drawing;
using OfficeOpenXml.Style;
using System.IO.Compression;

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

    [Test]
    public static void FontNameColorBoldItalicStrike()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            var runs = new XlsxRichTextRun[]
            {
                new("Red", XlsxFont.Default.WithName("Times New Roman").With(Color.Red).WithSize(14)),
                new(" Bold", XlsxFont.Default.WithBold()),
                new(" Italic", XlsxFont.Default.WithItalic()),
                new(" Strike", XlsxFont.Default.WithStrike()),
                new(" Normal", XlsxFont.Default)
            };

            xlsxWriter
                .BeginWorksheet("Sheet1")
                .BeginRow()
                .WriteRichText(runs, XlsxStyle.Default);
        }

        using var package = new ExcelPackage(stream);
        var cell = package.Workbook.Worksheets[0].Cells["A1"];
        cell.IsRichText.ShouldBeTrue();
        cell.Value.ShouldBe("Red Bold Italic Strike Normal");
        cell.RichText.Count.ShouldBe(5);

        cell.RichText[0].FontName.ShouldBe("Times New Roman");
        AreColorComponentsEqual(cell.RichText[0].Color, Color.Red).ShouldBeTrue();
        cell.RichText[0].Size.ShouldBe(14);

        cell.RichText[1].Bold.ShouldBeTrue();
        cell.RichText[2].Italic.ShouldBeTrue();
        cell.RichText[3].Strike.ShouldBeTrue();
        
        cell.RichText[4].Bold.ShouldBeFalse();
        cell.RichText[4].Italic.ShouldBeFalse();
        cell.RichText[4].Strike.ShouldBeFalse();
    }

    private static bool AreColorComponentsEqual(Color c1, Color c2)
    {
        return c1.A == c2.A && c1.R == c2.R && c1.G == c2.G && c1.B == c2.B;
    }

    [Test]
    public static void Underline_Types()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            var runs = new XlsxRichTextRun[]
            {
                new("None", XlsxFont.Default.WithUnderline(XlsxFont.Underline.None)),
                new(" Single", XlsxFont.Default.WithUnderline(XlsxFont.Underline.Single)),
                new(" Double", XlsxFont.Default.WithUnderline(XlsxFont.Underline.Double)),
                new(" SingleAcc", XlsxFont.Default.WithUnderline(XlsxFont.Underline.SingleAccounting)),
                new(" DoubleAcc", XlsxFont.Default.WithUnderline(XlsxFont.Underline.DoubleAccounting))
            };

            xlsxWriter
                .BeginWorksheet("Sheet1")
                .BeginRow()
                .WriteRichText(runs, XlsxStyle.Default);
        }

        // Validate underline types by inspecting the generated XML
        var bytes = stream.ToArray();
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        var sheetEntry = zip.GetEntry("xl/worksheets/sheet1.xml");
        sheetEntry.ShouldNotBeNull();
        using var sheetStream = sheetEntry.Open();
        using var reader = new StreamReader(sheetStream, Encoding.UTF8);
        var xml = reader.ReadToEnd();

        // Single underline => <u/>
        xml.ShouldContain("<u/>");
        // Double => <u val=\"double\"/>
        xml.ShouldContain("<u val=\"double\"/>");
        // SingleAccounting
        xml.ShouldContain("<u val=\"singleAccounting\"/>");
        // DoubleAccounting
        xml.ShouldContain("<u val=\"doubleAccounting\"/>");
    }

    [Test]
    public static void Whitespace_Preserve_PerRun()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            var runs = new XlsxRichTextRun[]
            {
                new(" lead", XlsxFont.Default),
                new("middle  spaces", XlsxFont.Default),
                new("trail ", XlsxFont.Default)
            };
            xlsxWriter.BeginWorksheet("Sheet1").BeginRow().WriteRichText(runs, XlsxStyle.Default);
        }
        using var package = new ExcelPackage(stream);
        var cell = package.Workbook.Worksheets[0].Cells["A1"];
        cell.IsRichText.ShouldBeTrue();
        cell.Value.ShouldBe(" leadmiddle  spacestrail ");
        cell.RichText[0].Text.ShouldBe(" lead");
        cell.RichText[1].Text.ShouldBe("middle  spaces");
        cell.RichText[2].Text.ShouldBe("trail ");
    }

    [Test]
    public static void NullOrEmptyRuns_WriteEmptyCell()
    {
        // Null runs
        using (var stream1 = new MemoryStream())
        {
            using (var xlsxWriter = new XlsxWriter(stream1))
                xlsxWriter.BeginWorksheet("S").BeginRow().WriteRichText(null, XlsxStyle.Default);
            using var package1 = new ExcelPackage(stream1);
            var cell1 = package1.Workbook.Worksheets[0].Cells["A1"];
            cell1.IsRichText.ShouldBeFalse();
            cell1.Value.ShouldBeNull();
        }

        // Empty list
        using (var stream2 = new MemoryStream())
        {
            using (var xlsxWriter = new XlsxWriter(stream2))
                xlsxWriter.BeginWorksheet("S").BeginRow().WriteRichText(Array.Empty<XlsxRichTextRun>(), XlsxStyle.Default);
            using var package2 = new ExcelPackage(stream2);
            var cell2 = package2.Workbook.Worksheets[0].Cells["A1"];
            cell2.IsRichText.ShouldBeFalse();
            cell2.Value.ShouldBeNull();
        }
    }

    [Test]
    public static void NullRunObjects_Ignored_And_EmptyRunText_Kept()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            IEnumerable<XlsxRichTextRun> runs = new object?[]
            {
                new XlsxRichTextRun("Hello", XlsxFont.Default),
                null,
                new XlsxRichTextRun("", XlsxFont.Default.WithBold())
            }.Cast<XlsxRichTextRun>();
            xlsxWriter.BeginWorksheet("S").BeginRow().WriteRichText(runs, XlsxStyle.Default);
        }
        using var package = new ExcelPackage(stream);
        var cell = package.Workbook.Worksheets[0].Cells["A1"];
        cell.IsRichText.ShouldBeTrue();
        cell.Value.ShouldBe("Hello");
        cell.RichText.Count.ShouldBe(2); // null run ignored, empty run kept
        cell.RichText[0].Text.ShouldBe("Hello");
        cell.RichText[1].Text.ShouldBe("");
        cell.RichText[1].Bold.ShouldBeTrue();
    }

    [Test]
    public static void DefaultFontUsed_WhenRunFontIsNull()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            var runs = new XlsxRichTextRun[] { new("abc") };
            xlsxWriter.BeginWorksheet("S").BeginRow().WriteRichText(runs, XlsxStyle.Default);
        }
        using var package = new ExcelPackage(stream);
        var rt = package.Workbook.Worksheets[0].Cells["A1"].RichText[0];
        rt.FontName.ShouldBeEmpty();
    }

    [Test]
    public static void MergedCells_WithRichText_ColumnSpan()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            var runs = new XlsxRichTextRun[] { new("Merged", XlsxFont.Default.WithBold()) };
            xlsxWriter.BeginWorksheet("S").BeginRow().WriteRichText(runs, XlsxStyle.Default, columnSpan: 3);
        }
        using var package = new ExcelPackage(stream);
        var ws = package.Workbook.Worksheets[0];
        ws.MergedCells.Count.ShouldBe(1);
        ws.MergedCells[0].ShouldBe("A1:C1");
        var cell = ws.Cells["A1"];
        cell.IsRichText.ShouldBeTrue();
        cell.RichText[0].Text.ShouldBe("Merged");
        cell.RichText[0].Bold.ShouldBeTrue();
    }

    [Test]
    public static void InvalidCharacters_RunText_ThrowOrSkip()
    {
        // Throw when skipInvalidCharacters=false (default)
        using (var stream1 = new MemoryStream())
        {
            Should.Throw<System.Xml.XmlException>(() =>
            {
                using var xlsxWriter = new XlsxWriter(stream1);
                var runs = new[] { new XlsxRichTextRun("a\0b", XlsxFont.Default) };
                xlsxWriter.BeginWorksheet("S").BeginRow().WriteRichText(runs, XlsxStyle.Default);
            });
        }

        // Skip when skipInvalidCharacters=true
        using (var stream2 = new MemoryStream())
        {
            using (var xlsxWriter = new XlsxWriter(stream2, skipInvalidCharacters: true))
            {
                var runs = new[] { new XlsxRichTextRun("a\0b", XlsxFont.Default) };
                xlsxWriter.BeginWorksheet("S").BeginRow().WriteRichText(runs, XlsxStyle.Default);
            }
            using var package2 = new ExcelPackage(stream2);
            var cell2 = package2.Workbook.Worksheets[0].Cells["A1"];
            cell2.Value.ShouldBe("ab");
            cell2.RichText[0].Text.ShouldBe("ab");
        }
    }

    [Test]
    public static void InvalidCharacters_FontName_ThrowOrSkip()
    {
        // Throw
        using (var stream1 = new MemoryStream())
        {
            Should.Throw<System.Xml.XmlException>(() =>
            {
                using var xlsxWriter = new XlsxWriter(stream1);
                var runs = new[] { new XlsxRichTextRun("X", XlsxFont.Default.WithName("A\0B")) };
                xlsxWriter.BeginWorksheet("S").BeginRow().WriteRichText(runs, XlsxStyle.Default);
            });
        }

        // Skip
        using (var stream2 = new MemoryStream())
        {
            using (var xlsxWriter = new XlsxWriter(stream2, skipInvalidCharacters: true))
            {
                var runs = new[] { new XlsxRichTextRun("X", XlsxFont.Default.WithName("A\0B")) };
                xlsxWriter.BeginWorksheet("S").BeginRow().WriteRichText(runs, XlsxStyle.Default);
            }
            using var package2 = new ExcelPackage(stream2);
            var rt = package2.Workbook.Worksheets[0].Cells["A1"].RichText[0];
            rt.FontName.ShouldBe("AB");
        }
    }
}