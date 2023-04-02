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
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using LargeXlsx;

namespace Examples;

public static class StyledLargeCreateStyles
{
    private const int RowCount = 50000;
    private const int ColumnCount = 180;
    private const int ColorCount = 100;

    public static void Run()
    {
        var stopwatch = Stopwatch.StartNew();
        DoRun();
        stopwatch.Stop();
        Console.WriteLine($"{nameof(StyledLargeCreateStyles)} completed {RowCount} rows, {ColumnCount} columns and {ColorCount} colors in {stopwatch.ElapsedMilliseconds} ms.");
    }

    private static void DoRun()
    {
        var rnd = new Random();
        using var stream = new FileStream($"{nameof(StyledLargeCreateStyles)}.xlsx", FileMode.Create, FileAccess.Write);
        using var xlsxWriter = new XlsxWriter(stream);
        var colors = Enumerable.Repeat(0, 100).Select(_ => Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256))).ToList();
        var headerStyle = new XlsxStyle(
            new XlsxFont("Calibri", 10.5, Color.White, bold: true),
            new XlsxFill(Color.FromArgb(0, 0x45, 0x86)),
            XlsxBorder.None,
            XlsxNumberFormat.General,
            XlsxAlignment.Default);

        xlsxWriter.BeginWorksheet("Sheet1", 1, 1);
        xlsxWriter.BeginRow();
        for (var j = 0; j < ColumnCount; j++)
            xlsxWriter.Write($"Column {j}", headerStyle);
        var colorIndex = 0;
        for (var i = 0; i < RowCount; i++)
        {
            xlsxWriter.BeginRow().Write($"Row {i}");
            for (var j = 1; j < 180; j++)
            {
                xlsxWriter.Write(i * ColumnCount + j, XlsxStyle.Default.With(new XlsxFill(colors[colorIndex])));
                colorIndex = (colorIndex + 1) % colors.Count;
            }
        }
    }
}