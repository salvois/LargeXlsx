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
namespace LargeXlsx
{
    public class XlsxFont
    {
        public static readonly XlsxFont Default = new XlsxFont(0, "Calibri", 11, "000000");
        internal const int FirstAvailableId = 1;

        public int Id { get; }
        public string FontName { get; }
        public double FontSize { get; }
        public string HexRgbColor { get; }
        public bool Bold { get; }
        public bool Italic { get; }
        public bool Strike { get; }

        internal XlsxFont(int id, string fontName, double fontSize, string hexRgbColor, bool bold = false, bool italic = false, bool strike = false)
        {
            Id = id;
            FontName = fontName;
            FontSize = fontSize;
            HexRgbColor = hexRgbColor;
            Bold = bold;
            Italic = italic;
            Strike = strike;
        }
    }
}