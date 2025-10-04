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
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;

namespace Examples;

public static class LargeAsync
{
    private const int RowCount = 50000;
    private const int ColumnCount = 180;

    public static async Task Run()
    {
        var stopwatch = Stopwatch.StartNew();
        var bufferCapacity = await DoRun(requireCellReferences: true);
        stopwatch.Stop();
        Console.WriteLine($"{nameof(LargeAsync),20} requiring references\t{RowCount}x{ColumnCount}\tBuffer capacity: {bufferCapacity}\tElapsed ms: {stopwatch.ElapsedMilliseconds}");

        stopwatch.Restart();
        bufferCapacity = await DoRun(requireCellReferences: false);
        stopwatch.Stop();
        Console.WriteLine($"{nameof(LargeAsync),20} omitting references\t{RowCount}x{ColumnCount}\tBuffer capacity: {bufferCapacity}\tElapsed ms: {stopwatch.ElapsedMilliseconds}");
    }

    private static async Task<int> DoRun(bool requireCellReferences)
    {
#if NETCOREAPP2_1_OR_GREATER
        await using var stream = new FileStream($"{nameof(LargeAsync)}_{requireCellReferences}.xlsx", FileMode.Create, FileAccess.Write);
#else
        using var stream = new FileStream($"{nameof(LargeAsync)}_{requireCellReferences}.xlsx", FileMode.Create, FileAccess.Write);
#endif
        await using var xlsxWriter = new XlsxWriter(stream, requireCellReferences: requireCellReferences);
        var whiteFont = new XlsxFont("Calibri", 11, Color.White, bold: true);
        var blueFill = new XlsxFill(Color.FromArgb(0, 0x45, 0x86));
        var headerStyle = new XlsxStyle(whiteFont, blueFill, XlsxBorder.None, XlsxNumberFormat.General, XlsxAlignment.Default);
        var numberStyle = XlsxStyle.Default.With(XlsxNumberFormat.ThousandTwoDecimal);

        await xlsxWriter.BeginWorksheetAsync("Sheet1", 1, 1);
        await xlsxWriter.BeginRowAsync();
        for (var j = 0; j < ColumnCount; j++)
            xlsxWriter.Write($"Column {j}", headerStyle);
        for (var i = 0; i < RowCount; i++)
        {
            await xlsxWriter.BeginRowAsync();
            xlsxWriter.Write($"Row {i}");
            for (var j = 1; j < ColumnCount; j++)
                xlsxWriter.Write(i * 1000 + j, numberStyle);
        }
        return xlsxWriter.BufferCapacity;
    }
}