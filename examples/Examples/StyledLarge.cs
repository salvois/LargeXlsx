/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020 Salvatore ISAJA. All rights reserved.

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

namespace Examples
{
    public static class StyledLarge
    {
        public static void Run()
        {
            var rnd = new Random();
            var stopwatch = Stopwatch.StartNew();
            using (var stream = new FileStream($"{nameof(StyledLarge)}.xlsx", FileMode.Create, FileAccess.Write))
            using (var xlsxWriter = new XlsxWriter(stream))
            {
                var whiteFont = new XlsxFont("Calibri", 11, Color.White, bold: true);
                var blueFill = new XlsxFill(XlsxFill.Pattern.Solid, Color.FromArgb(0, 0x45, 0x86));
                var headerStyle = new XlsxStyle(whiteFont, blueFill, XlsxBorder.None, XlsxNumberFormat.General);
                var cellStyles = Enumerable.Repeat(0, 100)
                    .Select(_ =>
                    {
                        var color = Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256));
                        var fill = new XlsxFill(XlsxFill.Pattern.Solid, color);
                        return new XlsxStyle(XlsxFont.Default, fill, XlsxBorder.None, XlsxNumberFormat.General);
                    })
                    .ToList();

                xlsxWriter.BeginWorksheet("Sheet1", 1, 1);
                xlsxWriter.BeginRow();
                for (var j = 0; j < 180; j++)
                    xlsxWriter.Write($"Column {j}", headerStyle);
                var cellStyleIndex = 0;
                for (var i = 0; i < 50000; i++)
                {
                    xlsxWriter.BeginRow().Write($"Row {i}");
                    for (var j = 1; j < 180; j++)
                    {
                        xlsxWriter.Write(i * 1000 + j, cellStyles[cellStyleIndex]);
                        cellStyleIndex = (cellStyleIndex + 1) % cellStyles.Count;
                    }
                }
            }
            stopwatch.Stop();
            Console.WriteLine($"Completed in {stopwatch.ElapsedMilliseconds} ms. Press any key...");
            Console.ReadKey();
        }
    }
}