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
using System.Drawing;
using System.IO;
using LargeXlsx;

namespace Examples
{
    public static class Alignment
    {
        public static void Run()
        {
            using (var stream = new FileStream($"{nameof(Alignment)}.xlsx", FileMode.Create, FileAccess.Write))
            using (var xlsxWriter = new XlsxWriter(stream))
            {
                xlsxWriter
                    .BeginWorksheet("Sheet 1", columns: new[] { XlsxColumn.Formatted(width: 40) })
                    .BeginRow().Write("Left", XlsxStyle.Default.With(new XlsxAlignment(horizontal: XlsxAlignment.Horizontal.Left)))
                    .BeginRow().Write("Center", XlsxStyle.Default.With(new XlsxAlignment(horizontal: XlsxAlignment.Horizontal.Center)))
                    .BeginRow().Write("Right", XlsxStyle.Default.With(new XlsxAlignment(horizontal: XlsxAlignment.Horizontal.Right)))
                    .BeginRow(height: 30).Write("Top", XlsxStyle.Default.With(new XlsxAlignment(vertical: XlsxAlignment.Vertical.Top)))
                    .BeginRow(height: 30).Write("Middle", XlsxStyle.Default.With(new XlsxAlignment(vertical: XlsxAlignment.Vertical.Center)))
                    .BeginRow(height: 30).Write("Bottom", XlsxStyle.Default.With(new XlsxAlignment(vertical: XlsxAlignment.Vertical.Bottom)))
                    .BeginRow(height: 90).Write("Rotated by 45°", XlsxStyle.Default.With(new XlsxAlignment(textRotation: 45)))
                    .BeginRow(height: 120).Write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt" +
                                                " ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation" +
                                                " ullamco laboris nisi ut aliquip ex ea commodo consequat.",
                        XlsxStyle.Default.With(new XlsxAlignment(horizontal: XlsxAlignment.Horizontal.Justify, vertical: XlsxAlignment.Vertical.Justify, wrapText: true)))
                    .BeginRow().Write("Lorem ipsum dolor sit amet, consectetur adipiscing elit", XlsxStyle.Default.With(new XlsxAlignment(shrinkToFit: true)));
            }
        }

        private static XlsxStyle With(this XlsxStyle style, XlsxAlignment alignment) =>
            new XlsxStyle(style.Font, style.Fill, style.Border, style.NumberFormat, alignment);
    }
}