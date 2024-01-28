/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2024 Salvatore ISAJA. All rights reserved.

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
    public class XlsxColumn
    {
        public int Count { get; }
        public bool Hidden { get; }
        public XlsxStyle Style { get; }
        public double? Width { get; }

        public static XlsxColumn Unformatted(int count = 1)
        {
            return new XlsxColumn(count, false, null, null);
        }

        public static XlsxColumn Formatted(double width, int count = 1, bool hidden = false, XlsxStyle style = null)
        {
            return new XlsxColumn(count, hidden, style, width);
        }

        private XlsxColumn(int count, bool hidden, XlsxStyle style, double? width)
        {
            Count = count;
            Hidden = hidden;
            Style = style;
            Width = width;
        }
    }
}