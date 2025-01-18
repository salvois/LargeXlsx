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
using LargeXlsx;

namespace Examples;

public static class RowFormatting
{
    public static void Run()
    {
        using var stream = new FileStream($"{nameof(RowFormatting)}.xlsx", FileMode.Create, FileAccess.Write);
        using var xlsxWriter = new XlsxWriter(stream);
        var blueStyle = new XlsxStyle(XlsxFont.Default.With(Color.White), new XlsxFill(Color.FromArgb(0, 0x45, 0x86)), XlsxBorder.None, XlsxNumberFormat.General, XlsxAlignment.Default);
        xlsxWriter
            .BeginWorksheet("Sheet 1")
            .BeginRow().Write("A1").Write("B1").Write("C1")
            .BeginRow(hidden: true).Write("A2").Write("B2").Write("C2")
            .BeginRow().Write("A3").Write("B3").Write("C3")
            .BeginRow(height: 36.5).Write("A4").Write("B4").Write("C4")
            .BeginRow(style: blueStyle).Write("A5").SkipColumns(1).Write("C5");
    }
}