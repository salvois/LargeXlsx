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
using System.Drawing;

namespace LargeXlsx
{
    public class XlsxFill : IEquatable<XlsxFill>
    {
        public enum Pattern
        {
            None,
            Gray125,
            Solid
        }

        public static readonly XlsxFill None = new XlsxFill(Pattern.None, Color.White);
        public static readonly XlsxFill Gray125 = new XlsxFill(Pattern.Gray125, Color.White);

        public Pattern PatternType { get; }
        public Color Color { get; }

        public XlsxFill(Pattern patternType, Color color)
        {
            PatternType = patternType;
            Color = color;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as XlsxFill);
        }

        public bool Equals(XlsxFill other)
        {
            return other != null && PatternType == other.PatternType && Color == other.Color;
        }

        public override int GetHashCode()
        {
            var hashCode = 493172489;
            hashCode = hashCode * -1521134295 + PatternType.GetHashCode();
            hashCode = hashCode * -1521134295 + Color.GetHashCode();
            return hashCode;
        }

        public static bool operator ==(XlsxFill fill1, XlsxFill fill2)
        {
            return EqualityComparer<XlsxFill>.Default.Equals(fill1, fill2);
        }

        public static bool operator !=(XlsxFill fill1, XlsxFill fill2)
        {
            return !(fill1 == fill2);
        }
    }
}