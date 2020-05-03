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
using System.Drawing;
using System.IO;
using LargeXlsx;

namespace Examples
{
    public static class Columns
    {
        public static void Run()
        {
            var rnd = new Random();
            using (var stream = new FileStream($"{nameof(Columns)}.xlsx", FileMode.Create, FileAccess.Write))
            using (var xlsxWriter = new XlsxWriter(stream))
            {
                var blueStyle = new XlsxStyle(
                    new XlsxFont(XlsxFont.Default.FontName, XlsxFont.Default.FontSize, Color.White),
                    new XlsxFill(XlsxFill.Pattern.Solid, Color.FromArgb(0, 0x45, 0x86)),
                    XlsxBorder.None,
                    XlsxNumberFormat.General);

                xlsxWriter
                    .BeginWorksheet("Sheet 1", columns: new[]
                    {
                        XlsxColumn.Formatted(count: 2, width: 20),
                        XlsxColumn.Unformatted(3),
                        XlsxColumn.Formatted(style: blueStyle, width: 9),
                        XlsxColumn.Formatted(hidden: true, width: 0)
                    });
                for (var i = 0; i < 10; i++)
                {
                    xlsxWriter.BeginRow();
                    for (var j = 0; j < 10; j++)
                        xlsxWriter.Write(rnd.Next());
                }
            }
        }
    }
}