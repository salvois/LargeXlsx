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
using System.Collections.Generic;

namespace LargeXlsx
{
    public class XlsxStyle : IEquatable<XlsxStyle>
    {
        public static readonly XlsxStyle Default = new XlsxStyle(XlsxFont.Default, XlsxFill.None, XlsxBorder.None, XlsxNumberFormat.General);

        public XlsxFont Font { get; }
        public XlsxFill Fill { get; }
        public XlsxBorder Border { get; }
        public XlsxNumberFormat NumberFormat { get; }

        public XlsxStyle(XlsxFont font, XlsxFill fill, XlsxBorder border, XlsxNumberFormat numberFormat)
        {
            Font = font;
            Fill = fill;
            Border = border;
            NumberFormat = numberFormat;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as XlsxStyle);
        }

        public bool Equals(XlsxStyle other)
        {
            return other != null && Font.Equals(other.Font) && Fill.Equals(Fill) && Border.Equals(other.Border) && NumberFormat.Equals(NumberFormat);
        }

        public override int GetHashCode()
        {
            var hashCode = 428549002;
            hashCode = hashCode * -1521134295 + Font.GetHashCode();
            hashCode = hashCode * -1521134295 + Fill.GetHashCode();
            hashCode = hashCode * -1521134295 + Border.GetHashCode();
            hashCode = hashCode * -1521134295 + NumberFormat.GetHashCode();
            return hashCode;
        }

        public static bool operator ==(XlsxStyle style1, XlsxStyle style2)
        {
            return EqualityComparer<XlsxStyle>.Default.Equals(style1, style2);
        }

        public static bool operator !=(XlsxStyle style1, XlsxStyle style2)
        {
            return !(style1 == style2);
        }
    }
}