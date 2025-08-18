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

namespace Examples
{
    public static class HideGridlines
    {
        public static void Run()
        {
            var xlsxStyle =
                XlsxStyle.Default.With(XlsxBorder.Around(new XlsxBorder.Line(Color.Black, XlsxBorder.Style.Dotted)));
            using (var stream = new FileStream($"{nameof(HideGridlines)}.xlsx", FileMode.Create, FileAccess.Write))
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                {
                    xlsxWriter
                        .BeginWorksheet("HiddenGridlines", showGridLines: false)
                        .BeginRow().Write("A1", xlsxStyle).Write("B1", xlsxStyle).Write("C1", xlsxStyle)
                        .BeginRow().Write("A2", xlsxStyle).Write("B2", xlsxStyle).Write("C2", xlsxStyle)
                        .BeginRow().Write("A3", xlsxStyle).Write("B3", xlsxStyle).Write("C3", xlsxStyle)
                        .BeginWorksheet("HiddenHeaders", showHeaders: false)
                        .BeginRow()
                        .BeginRow().Write("").Write("This sheet have row and columns headers hidden by default")
                        .BeginRow();
                }
            }
        }
    }
}