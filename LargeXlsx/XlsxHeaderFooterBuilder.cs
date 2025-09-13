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
using System.Text;

namespace LargeXlsx
{
    public class XlsxHeaderFooterBuilder
    {
        public override string ToString() => _sb.ToString();

        public XlsxHeaderFooterBuilder Left() => DoAppend("&L");
        public XlsxHeaderFooterBuilder Center() => DoAppend("&C");
        public XlsxHeaderFooterBuilder Right() => DoAppend("&R");
        public XlsxHeaderFooterBuilder Text(string text) => DoAppend(text.Replace("&", "&&"));
        public XlsxHeaderFooterBuilder CurrentDate() => DoAppend("&D");
        public XlsxHeaderFooterBuilder CurrentTime() => DoAppend("&T");
        public XlsxHeaderFooterBuilder FileName() => DoAppend("&F");
        public XlsxHeaderFooterBuilder FilePath() => DoAppend("&Z");
        public XlsxHeaderFooterBuilder NumberOfPages() => DoAppend("&N");
        public XlsxHeaderFooterBuilder PageNumber(int offset = 0) => offset == 0 ? DoAppend("&P") : DoAppend($"&P{offset:+0;-0}");
        public XlsxHeaderFooterBuilder SheetName() => DoAppend("&A");
        public XlsxHeaderFooterBuilder FontSize(int points) => DoAppend($"&{points:0}");
        public XlsxHeaderFooterBuilder Font(string name, bool bold = false, bool italic = false) => DoAppend($"&\"{name},{GetFontType(bold, italic)}\"");
        public XlsxHeaderFooterBuilder Font(bool bold = false, bool italic = false) => DoAppend($"&\"-,{GetFontType(bold, italic)}\"");
        public XlsxHeaderFooterBuilder Bold() => DoAppend("&B");
        public XlsxHeaderFooterBuilder Italic() => DoAppend("&I");
        public XlsxHeaderFooterBuilder Underline() => DoAppend("&U");
        public XlsxHeaderFooterBuilder DoubleUnderline() => DoAppend("&E");
        public XlsxHeaderFooterBuilder StrikeThrough() => DoAppend("&S");
        public XlsxHeaderFooterBuilder Subscript() => DoAppend("&Y");
        public XlsxHeaderFooterBuilder Superscript() => DoAppend("&X");

        private readonly StringBuilder _sb = new StringBuilder();

        private XlsxHeaderFooterBuilder DoAppend(string text)
        {
            _sb.Append(text);
            return this;
        }

        private static string GetFontType(bool bold, bool italic)
        {
            if (!bold && !italic) return "Regular";
            if (bold && !italic) return "Bold";
            if (!bold) return "Italic";
            return "Bold Italic";
        }
    }
}