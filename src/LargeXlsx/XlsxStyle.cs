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
    public struct XlsxStyle
    {
        public static readonly XlsxStyle Default = new XlsxStyle(0, XlsxFont.Default, XlsxFill.None, XlsxBorder.None, XlsxNumberFormat.General);
        internal const int FirstAvailableId = 1;

        public int Id { get; }
        public XlsxFont Font { get; }
        public XlsxFill Fill { get; }
        public XlsxBorder Border { get; }
        public XlsxNumberFormat NumberFormat { get; }

        internal XlsxStyle(int id, XlsxFont font, XlsxFill fill, XlsxBorder border, XlsxNumberFormat numberFormat)
        {
            Id = id;
            Font = font;
            Fill = fill;
            Border = border;
            NumberFormat = numberFormat;
        }
    }
}