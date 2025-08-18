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
using Shouldly;

namespace LargeXlsx.Tests
{

    [TestFixture]
    public static class XlsxHeaderFooterBuilderTest
    {
        [Test]
        public static void Simple() =>
            new XlsxHeaderFooterBuilder()
                .Left().Bold().Italic().Underline().Text("Left&Formatted")
                .Center().DoubleUnderline().Subscript().PageNumber().NumberOfPages().CurrentDate().CurrentTime()
                .FilePath().FileName()
                .Right().Superscript().StrikeThrough().SheetName()
                .ToString()
                .ShouldBe("&L&B&I&ULeft&&Formatted"
                          + "&C&E&Y&P&N&D&T&Z&F"
                          + "&R&X&S&A");

        [TestCase(0, "&P")]
        [TestCase(1, "&P+1")]
        [TestCase(-1, "&P-1")]
        [TestCase(42, "&P+42")]
        [TestCase(-69, "&P-69")]
        public static void PageNumber(int offset, string expected) =>
            new XlsxHeaderFooterBuilder().PageNumber(offset).ToString().ShouldBe(expected);

        [TestCase(1, "&1")]
        [TestCase(42, "&42")]
        public static void FontSize(int points, string expected) =>
            new XlsxHeaderFooterBuilder().FontSize(points).ToString().ShouldBe(expected);

        [TestCase("Times New Roman", false, false, "&\"Times New Roman,Regular\"")]
        [TestCase("Times New Roman", true, false, "&\"Times New Roman,Bold\"")]
        [TestCase("Times New Roman", false, true, "&\"Times New Roman,Italic\"")]
        [TestCase("Times New Roman", true, true, "&\"Times New Roman,Bold Italic\"")]
        public static void Font(string name, bool bold, bool italic, string expected) =>
            new XlsxHeaderFooterBuilder().Font(name, bold, italic).ToString().ShouldBe(expected);

        [TestCase(false, false, "&\"-,Regular\"")]
        [TestCase(true, false, "&\"-,Bold\"")]
        [TestCase(false, true, "&\"-,Italic\"")]
        [TestCase(true, true, "&\"-,Bold Italic\"")]
        public static void FontWithoutName(bool bold, bool italic, string expected) =>
            new XlsxHeaderFooterBuilder().Font(bold, italic).ToString().ShouldBe(expected);
    }
}