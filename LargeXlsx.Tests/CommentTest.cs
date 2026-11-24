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
}
