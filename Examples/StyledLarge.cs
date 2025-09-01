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
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using LargeXlsx;
using SharpCompress.Compressors.Deflate;

namespace Examples;

public static class StyledLarge
{
    private const int RowCount = 50000;
    private const int ColumnCount = 180;
    private const int ColorCount = 100;

    public static void Run()
    {
        var stopwatch = Stopwatch.StartNew();
        DoRun(requireCellReferences: true);
        stopwatch.Stop();
        Console.WriteLine($"{nameof(StyledLarge)} requiring references completed {RowCount} rows, {ColumnCount} columns and {ColorCount} colors in {stopwatch.ElapsedMilliseconds} ms.");

        stopwatch.Restart();
        DoRun(requireCellReferences: false);
        stopwatch.Stop();
        Console.WriteLine($"{nameof(StyledLarge)} omitting references completed {RowCount} rows, {ColumnCount} columns and {ColorCount} colors in {stopwatch.ElapsedMilliseconds} ms.");
    }

    private static void DoRun(bool requireCellReferences)
    {
        var rnd = new Random();
        using var stream = new FileStream($"{nameof(StyledLarge)}_{requireCellReferences}.xlsx", FileMode.Create, FileAccess.Write);
        using var xlsxWriter = new XlsxWriter(stream, compressionLevel: CompressionLevel.Level3, requireCellReferences: requireCellReferences);
        var headerStyle = new XlsxStyle(
            new XlsxFont("Calibri", 11, Color.White, bold: true),
            new XlsxFill(Color.FromArgb(0, 0x45, 0x86)),
            XlsxBorder.None,
            XlsxNumberFormat.General,
            XlsxAlignment.Default);
        var cellStyles = Enumerable.Repeat(0, 100)
            .Select(_ => XlsxStyle.Default.With(new XlsxFill(Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256)))))
            .ToList();

        xlsxWriter.BeginWorksheet("Sheet1", 1, 1);
        xlsxWriter.BeginRow();
        for (var j = 0; j < ColumnCount; j++)
            xlsxWriter.Write($"Column {j}", headerStyle);
        var cellStyleIndex = 0;
        for (var i = 0; i < RowCount; i++)
        {
            xlsxWriter.BeginRow().Write($"Row {i}");
            for (var j = 1; j < 180; j++)
            {
                xlsxWriter.Write(i * ColumnCount + j, cellStyles[cellStyleIndex]);
                cellStyleIndex = (cellStyleIndex + 1) % cellStyles.Count;
            }
            if (i % 100 == 0)
                xlsxWriter.Commit();
        }
    }
}