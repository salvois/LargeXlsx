/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2021 Salvatore ISAJA. All rights reserved.

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
using System.IO;
using System.Linq;
using LargeXlsx;
using SharpCompress.Compressors.Deflate;

namespace Examples
{
    public static class Zip64Huge
    {
        private const int RowCount = 1000000;
        private const int ColumnCount = 180;

        public static void Run()
        {
            var stopwatch = Stopwatch.StartNew();
            using (var stream = new FileStream($"{nameof(Zip64Huge)}.xlsx", FileMode.Create, FileAccess.Write))
            using (var xlsxWriter = new XlsxWriter(stream, compressionLevel: CompressionLevel.BestSpeed, useZip64: true))
            {
                xlsxWriter.BeginWorksheet("Sheet1", 1, 1);
                xlsxWriter.BeginRow();
                for (var j = 0; j < ColumnCount; j++)
                    xlsxWriter.Write($"Column {j}");
                for (var i = 0; i < RowCount; i++)
                {
                    xlsxWriter.BeginRow().Write($"Row {i}");
                    for (var j = 1; j < ColumnCount; j++)
                        xlsxWriter.Write(i * 100 + j);
                    if (i % 50000 == 0)
                        Console.WriteLine($"{nameof(Zip64Huge)} wrote {i} rows in {stopwatch.ElapsedMilliseconds} ms...");
                }
            }
            stopwatch.Stop();
            Console.WriteLine($"{nameof(Zip64Huge)} completed {RowCount} rows and {ColumnCount} columns in {stopwatch.ElapsedMilliseconds} ms.");
        }
    }
}