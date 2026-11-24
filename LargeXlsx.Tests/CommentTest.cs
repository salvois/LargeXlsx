/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2026 Salvatore ISAJA. All rights reserved.

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
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using Shouldly;

namespace LargeXlsx.Tests;

[TestFixture]
public static class CommentTest
{
    [Test]
    public static void NullTextThrows() =>
        Should.Throw<ArgumentNullException>(() => new XlsxComment(null!));

    [Test]
    public static void NonPositiveSizeThrows()
    {
        Should.Throw<ArgumentOutOfRangeException>(() => new XlsxComment("x", widthInColumns: 0));
        Should.Throw<ArgumentOutOfRangeException>(() => new XlsxComment("x", heightInRows: 0));
    }

    [Test]
    public static void DuplicateCellThrows()
    {
        using var stream = new MemoryStream();
        using var xlsxWriter = new XlsxWriter(stream);
        xlsxWriter.BeginWorksheet("Sheet1").BeginRow().AddComment(new XlsxComment("a"));
        Should.Throw<InvalidOperationException>(() => xlsxWriter.AddComment(1, 1, new XlsxComment("b")));
    }

    [Test]
    public static void CommentOutsideRowThrows()
    {
        using var stream = new MemoryStream();
        using var xlsxWriter = new XlsxWriter(stream);
        xlsxWriter.BeginWorksheet("Sheet1");
        Should.Throw<ArgumentOutOfRangeException>(() => xlsxWriter.AddComment(new XlsxComment("a")));
        Should.Throw<ArgumentOutOfRangeException>(() => xlsxWriter.AddComment(Limits.MaxRowCount + 1, 1, new XlsxComment("a")));
        Should.Throw<ArgumentOutOfRangeException>(() => xlsxWriter.AddComment(1, Limits.MaxColumnCount + 1, new XlsxComment("a")));
    }

    [Test]
    public static void CommentWithoutWorksheetThrows()
    {
        using var stream = new MemoryStream();
        using var xlsxWriter = new XlsxWriter(stream);
        Should.Throw<InvalidOperationException>(() => xlsxWriter.AddComment(1, 1, new XlsxComment("a")));
    }

    [Test]
    public static void SingleComment()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet1").BeginRow().Write("A1")
                .AddComment(new XlsxComment("Hello", author: "Alice")).Write("B1");

        using var package = EPPlusWrapper.Create(stream);
        var sheet = package.Workbook.Worksheets[0];
        sheet.Comments.Count.ShouldBe(1);
        sheet.Cells["B1"].Comment.Text.ShouldBe("Hello");
        sheet.Cells["B1"].Comment.Author.ShouldBe("Alice");
        sheet.Cells["B1"].Comment.Visible.ShouldBeFalse();
        sheet.Cells["B1"].Value.ShouldBe("B1");
    }

    [Test]
    public static void VisibleComment()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet1").BeginRow().AddComment(new XlsxComment("Hello", visible: true)).Write("A1");

        using var package = EPPlusWrapper.Create(stream);
        package.Workbook.Worksheets[0].Cells["A1"].Comment.Visible.ShouldBeTrue();
    }

    [Test]
    public static void NullAuthor()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet1").BeginRow().AddComment(new XlsxComment("Hello")).Write("A1");

        using var package = EPPlusWrapper.Create(stream);
        package.Workbook.Worksheets[0].Cells["A1"].Comment.Author.ShouldBe("");
    }

    [Test]
    public static void SharedAndDistinctAuthors()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet1").BeginRow()
                .AddComment(new XlsxComment("one", author: "A")).Write()
                .AddComment(new XlsxComment("two", author: "B")).Write()
                .AddComment(new XlsxComment("three", author: "A")).Write();

        using var package = EPPlusWrapper.Create(stream);
        var sheet = package.Workbook.Worksheets[0];
        sheet.Comments.Count.ShouldBe(3);
        sheet.Cells["A1"].Comment.Author.ShouldBe("A");
        sheet.Cells["B1"].Comment.Author.ShouldBe("B");
        sheet.Cells["C1"].Comment.Author.ShouldBe("A");
        sheet.Cells["C1"].Comment.Text.ShouldBe("three");
    }

    [Test]
    public static void ExplicitPositionOutOfOrder()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet1").BeginRow().Write("A1")
                .AddComment(5, 3, new XlsxComment("C5"))
                .AddComment(2, 1, new XlsxComment("A2"));

        using var package = EPPlusWrapper.Create(stream);
        var sheet = package.Workbook.Worksheets[0];
        sheet.Comments.Count.ShouldBe(2);
        sheet.Cells["C5"].Comment.Text.ShouldBe("C5");
        sheet.Cells["A2"].Comment.Text.ShouldBe("A2");
    }

    [Test]
    public static void SpecialCharacters()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet1").BeginRow()
                .AddComment(new XlsxComment("  <a & b>\n\"line 2\" ", author: "<O'Neil & co>")).Write();

        using var package = EPPlusWrapper.Create(stream);
        var comment = package.Workbook.Worksheets[0].Cells["A1"].Comment;
        comment.Text.ShouldBe("  <a & b>\n\"line 2\" ");
        comment.Author.ShouldBe("<O'Neil & co>");
    }

    [Test]
    public static void EscapedAndInvalidXmlCharacters()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = CreateOpenXmlCompatibleWriter(stream, skipInvalidCharacters: false))
            xlsxWriter.BeginWorksheet("Sheet1").BeginRow().AddComment(new XlsxComment("a\u0001b_x0041_")).Write();

        GetCommentTexts(stream).ShouldBe(["a_x0001_b_x005F_x0041_"]);
    }

    [Test]
    public static void SkippedInvalidXmlCharacters()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream, skipInvalidCharacters: true))
            xlsxWriter.BeginWorksheet("Sheet1").BeginRow().AddComment(new XlsxComment("a\u0001b", author: "c\u0001d")).Write();

        using var package = EPPlusWrapper.Create(stream);
        var comment = package.Workbook.Worksheets[0].Cells["A1"].Comment;
        comment.Text.ShouldBe("ab");
        comment.Author.ShouldBe("cd");
    }

    [Test]
    public static void NoCommentsProducesNoParts()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = CreateOpenXmlCompatibleWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet1").BeginRow().Write("A1");

        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = document.WorkbookPart!.WorksheetParts.Single();
        worksheetPart.WorksheetCommentsPart.ShouldBeNull();
        worksheetPart.VmlDrawingParts.ShouldBeEmpty();
        worksheetPart.Worksheet!.Descendants<LegacyDrawing>().ShouldBeEmpty();
    }

    [Test]
    public static async Task AsyncDispose()
    {
        using var stream = new MemoryStream();
        var xlsxWriter = new XlsxWriter(stream);
        await xlsxWriter.BeginWorksheetAsync("Sheet1");
        await xlsxWriter.BeginRowAsync();
        xlsxWriter.AddComment(new XlsxComment("Hello", author: "Alice")).Write("A1");
        await xlsxWriter.DisposeAsync();

        using var package = EPPlusWrapper.Create(stream);
        var comment = package.Workbook.Worksheets[0].Cells["A1"].Comment;
        comment.Text.ShouldBe("Hello");
        comment.Author.ShouldBe("Alice");
    }

    [Test]
    public static void AllocatorReservesConsecutiveBlocks()
    {
        var allocator = new VmlShapeIdAllocator();
        allocator.AllocateBlocks(1).ShouldBe(1);
        allocator.AllocateBlocks(1023).ShouldBe(2);
        allocator.AllocateBlocks(1024).ShouldBe(3);
        allocator.AllocateBlocks(1).ShouldBe(5);
    }

    [Test]
    public static void AnchorStaysInsideGrid()
    {
        WorksheetComments.ComputeAnchor(0, 0, 2, 4).ShouldBe("1, 15, 0, 10, 3, 15, 4, 4");
        WorksheetComments.ComputeAnchor(Limits.MaxRowCount - 1, Limits.MaxColumnCount - 1, 2, 4)
            .ShouldBe($"{Limits.MaxColumnCount - 3}, 15, {Limits.MaxRowCount - 5}, 10, {Limits.MaxColumnCount - 1}, 15, {Limits.MaxRowCount - 1}, 4");
    }

    [Test]
    public static void CommentsOnSeveralSheets()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = CreateOpenXmlCompatibleWriter(stream))
        {
            // More than one VML block of shape ids on the first sheet, to catch overlaps with the next sheets
            xlsxWriter.BeginWorksheet("Sheet1");
            for (var i = 1; i <= 1100; i++)
                xlsxWriter.BeginRow().AddComment(new XlsxComment($"1A{i}")).Write();
            xlsxWriter.BeginWorksheet("Sheet2").BeginRow().AddComment(new XlsxComment("2A1")).Write().AddComment(new XlsxComment("2B1")).Write();
            xlsxWriter.BeginWorksheet("Sheet3").BeginRow().Write();
        }

        using (var package = EPPlusWrapper.Create(stream))
        {
            package.Workbook.Worksheets[0].Comments.Count.ShouldBe(1100);
            package.Workbook.Worksheets[1].Comments.Count.ShouldBe(2);
            package.Workbook.Worksheets[1].Cells["B1"].Comment.Text.ShouldBe("2B1");
            package.Workbook.Worksheets[2].Comments.Count.ShouldBe(0);
        }
        GetVmlIdMaps(stream).ShouldBe(["1,2", "3"]);
        var shapeIds = GetVmlShapeIds(stream);
        shapeIds.Length.ShouldBe(1102);
        shapeIds.Distinct().Count().ShouldBe(1102);
    }

    [Test]
    public static void ManyCommentsOnOneSheet()
    {
        const int commentCount = 2500;
        using var stream = new MemoryStream();
        using (var xlsxWriter = CreateOpenXmlCompatibleWriter(stream))
        {
            xlsxWriter.BeginWorksheet("Sheet1");
            for (var i = 1; i <= commentCount; i++)
                xlsxWriter.BeginRow().AddComment(new XlsxComment($"Comment {i}")).Write(i);
        }

        using (var package = EPPlusWrapper.Create(stream))
        {
            var sheet = package.Workbook.Worksheets[0];
            sheet.Comments.Count.ShouldBe(commentCount);
            sheet.Cells[$"A{commentCount}"].Comment.Text.ShouldBe($"Comment {commentCount}");
        }
        GetVmlIdMaps(stream).ShouldBe(["1,2,3"]);
        var shapeIds = GetVmlShapeIds(stream);
        shapeIds.Length.ShouldBe(commentCount);
        shapeIds.Distinct().Count().ShouldBe(commentCount);
    }

    [TestCase(false)]
    [TestCase(true)]
    public static void PartsAreSchemaValid(bool withComments)
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = CreateOpenXmlCompatibleWriter(stream))
        {
            // Other post-sheetData elements, so that <legacyDrawing> is checked to follow them
            xlsxWriter.BeginWorksheet("Sheet1")
                .SetHeaderFooter(new XlsxHeaderFooter(oddHeader: "&CHeader"))
                .BeginRow().AddMergedCell(1, 2).Write("A1").Write()
                .AddDataValidation(XlsxDataValidation.List(["Lorem", "Ipsum"])).Write()
                .BeginRow().AddRowPageBreak().Write("A2");
            if (withComments)
                xlsxWriter.AddComment(1, 1, new XlsxComment("one", author: "Alice")).AddComment(2, 1, new XlsxComment(" two ", visible: true));
        }

        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = document.WorkbookPart!.WorksheetParts.Single();
        var validator = new OpenXmlValidator(FileFormatVersions.Office2010);
        validator.Validate(worksheetPart.Worksheet!).Select(e => e.Description).ShouldBeEmpty();
        if (withComments)
        {
            worksheetPart.WorksheetCommentsPart.ShouldNotBeNull();
            worksheetPart.VmlDrawingParts.Count().ShouldBe(1);
            worksheetPart.Worksheet!.Elements<LegacyDrawing>().Count().ShouldBe(1);
            validator.Validate(worksheetPart.WorksheetCommentsPart!.Comments!).Select(e => e.Description).ShouldBeEmpty();
        }
    }

    [Test]
    public static async Task CommentsAreReleasedOnceWritten()
    {
        // Comments of a closed worksheet must not stay in RAM until the whole workbook is disposed
        var comments = new WorksheetComments();
        comments.Add(1, 1, new XlsxComment("a"));
        comments.Add(2, 1, new XlsxComment("b"));

        await comments.WritePartsAsync(new MemoryZipWriter(), new CustomWriter(65536), sheetId: 1, firstVmlBlock: 1, skipInvalidCharacters: false);

        comments.Count.ShouldBe(0);
    }

    private sealed class MemoryZipWriter : IZipWriter
    {
        public Stream CreateEntry(string path) => new MemoryStream();
        public void Dispose() { }
    }

    private static string[] GetVmlContents(Stream stream)
    {
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!.WorksheetParts
            .SelectMany(wp => wp.VmlDrawingParts)
            .Select(vp =>
            {
                using var reader = new StreamReader(vp.GetStream());
                return reader.ReadToEnd();
            })
            .ToArray();
    }

    private static string[] GetVmlShapeIds(Stream stream) =>
        GetVmlContents(stream).SelectMany(vml => Regex.Matches(vml, "<v:shape id=\"(_x0000_s\\d+)\"").Cast<Match>().Select(m => m.Groups[1].Value)).ToArray();

    private static string[] GetVmlIdMaps(Stream stream) =>
        GetVmlContents(stream).SelectMany(vml => Regex.Matches(vml, "<o:idmap v:ext=\"edit\" data=\"([^\"]*)\"").Cast<Match>().Select(m => m.Groups[1].Value)).ToArray();

    private static XlsxWriter CreateOpenXmlCompatibleWriter(Stream stream, bool skipInvalidCharacters = false) =>
#if NETCOREAPP2_1_OR_GREATER
        new(stream, skipInvalidCharacters: skipInvalidCharacters);
#else
        // ZIP64 does not work with DocumentFormat.OpenXml on .NET Framework
        new(new SharpCompressZipWriter(stream, XlsxCompressionLevel.Fastest, useZip64: false), skipInvalidCharacters: skipInvalidCharacters);
#endif

    private static string[] GetCommentTexts(Stream stream)
    {
        using var document = SpreadsheetDocument.Open(stream, false);
        var commentsPart = document.WorkbookPart!.WorksheetParts.Single().WorksheetCommentsPart;
        commentsPart.ShouldNotBeNull();
        return commentsPart.Comments!.CommentList!.Elements<Comment>().Select(c => c.CommentText!.InnerText).ToArray();
    }
}
