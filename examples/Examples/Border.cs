/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2023 Salvatore ISAJA. All rights reserved.

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

public static class Border
{
    public static void Run()
    {
        using var stream = new FileStream($"{nameof(Border)}.xlsx", FileMode.Create, FileAccess.Write);
        using var xlsxWriter = new XlsxWriter(stream);
        var leftBorderStyle = XlsxStyle.Default.With(new XlsxBorder(left: new XlsxBorder.Line(Color.DeepPink, XlsxBorder.Style.Thin)));
        var allBorderStyle = XlsxStyle.Default.With(XlsxBorder.Around(new XlsxBorder.Line(Color.CornflowerBlue, XlsxBorder.Style.Dashed)));
        var diagonalBorderStyle = XlsxStyle.Default.With(
            new XlsxBorder(diagonal: new XlsxBorder.Line(Color.Red, XlsxBorder.Style.Dotted), diagonalDown: true, diagonalUp: true));

        xlsxWriter
            .BeginWorksheet("Sheet1")
            .SkipRows(1)
            .BeginRow(height: 50).SkipColumns(1).Write("B1", leftBorderStyle).SkipColumns(1).Write("D1", allBorderStyle).Write("E1", diagonalBorderStyle)
            .BeginRow().SkipColumns(1).Write("B2", leftBorderStyle).SkipColumns(1).Write("D2", allBorderStyle).Write("E2", diagonalBorderStyle)
            .BeginRow().SkipColumns(1).Write(leftBorderStyle).SkipColumns(1).Write(allBorderStyle).Write(diagonalBorderStyle);
    }
}