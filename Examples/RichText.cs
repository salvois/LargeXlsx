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
using LargeXlsx;
using System.Drawing;
using System.IO;

namespace Examples;

public static class RichText
{
    public static void Run()
    {
        using var stream = new FileStream($"{nameof(RichText)}.xlsx", FileMode.Create, FileAccess.Write);
        using var xlsxWriter = new XlsxWriter(stream);
        var headerStyle = new XlsxStyle(
            new XlsxFont("Segoe UI", 9, Color.White, bold: true),
            new XlsxFill(Color.FromArgb(0, 0x45, 0x86)),
            XlsxStyle.Default.Border,
            XlsxStyle.Default.NumberFormat,
            XlsxAlignment.Default);

        xlsxWriter
            .BeginWorksheet("Sheet 1", columns: [XlsxColumn.Formatted(width: 120)])
            .SetDefaultStyle(XlsxStyle.Default)
            .BeginRow()
            .Write("Rich Text Example", XlsxStyle.Default.With(XlsxFont.Default.WithBold().WithSize(16)))
            .BeginRow()
            .WriteRichText(
            [
                new XlsxRichTextRun("Normal Text ", XlsxFont.Default),
                new XlsxRichTextRun("and ", XlsxFont.Default),
                new XlsxRichTextRun("Large Text ", XlsxFont.Default.WithSize(16)),
                new XlsxRichTextRun("and ", XlsxFont.Default),
                new XlsxRichTextRun("Bold Text ", XlsxFont.Default.WithBold()),
                new XlsxRichTextRun("and ", XlsxFont.Default),
                new XlsxRichTextRun("Italic Text ", XlsxFont.Default.WithItalic()),
                new XlsxRichTextRun("and ", XlsxFont.Default),
                new XlsxRichTextRun("Underline Text ", XlsxFont.Default.WithUnderline()),
                new XlsxRichTextRun("and ", XlsxFont.Default),
                new XlsxRichTextRun("Strike Text ", XlsxFont.Default.WithStrike()),
                new XlsxRichTextRun("and ", XlsxFont.Default),
                new XlsxRichTextRun("Hyperlink", XlsxFont.Default.WithUnderline().With(Color.Blue))
            ]);
    }
}