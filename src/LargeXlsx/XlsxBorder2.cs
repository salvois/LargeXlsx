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

namespace LargeXlsx
{
    public class XlsxBorder2
    {
        public enum Style
        {
            None,
            Thin,
            Medium,
            Dashed,
            Dotted,
            Thick,
            Double,
            Hair,
            MediumDashed,
            DashDot,
            MediumDashDot,
            DashDotDot,
            MediumDashDotDot,
            SlantDashDot
        }

        public static readonly XlsxBorder2 None = new XlsxBorder2(0, "000000");
        internal const int FirstAvailableId = 1;

        public int Id { get; }
        public string HexRgbColor { get; }
        public Style Top { get; }
        public Style Right { get; }
        public Style Bottom { get; }
        public Style Left { get; }

        internal XlsxBorder2(int id, string hexRgbColor, Style top = Style.None, Style right = Style.None, Style bottom = Style.None, Style left = Style.None)
        {
            Id = id;
            HexRgbColor = hexRgbColor;
            Top = top;
            Right = right;
            Bottom = bottom;
            Left = left;
        }

        internal static string GetStyleAttributeValue(Style style)
        {
            switch (style)
            {
                case Style.None: return "none";
                case Style.Thin: return "thin";
                case Style.Medium: return "medium";
                case Style.Dashed: return "dashed";
                case Style.Dotted: return "dotted";
                case Style.Thick: return "thick";
                case Style.Double: return "double";
                case Style.Hair: return "hair";
                case Style.MediumDashed: return "mediumDashed";
                case Style.DashDot: return "dashDot";
                case Style.MediumDashDot: return "mediumDashDot";
                case Style.DashDotDot: return "dashDotDot";
                case Style.MediumDashDotDot: return "mediumDashDotDot";
                case Style.SlantDashDot: return "slantDashDot";
                default: throw new ArgumentOutOfRangeException();
            }
        }
    }
}