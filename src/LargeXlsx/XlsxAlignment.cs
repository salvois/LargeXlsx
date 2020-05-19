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
    public class XlsxAlignment : IEquatable<XlsxAlignment>
    {
        public static readonly XlsxAlignment Default = new XlsxAlignment();

        public Horizontal HorizontalType { get; }
        public Vertical VerticalType { get; }
        public int Indent { get; }
        public bool JustifyLastLine { get; }
        public ReadingOrder ReadingOrderType { get; }
        public bool ShrinkToFit { get; }
        public int TextRotation { get; }
        public bool WrapText { get; }

        public enum Horizontal
        {
            General,
            Left,
            Center,
            Right,
            Fill,
            Justify,
            CenterContinuous,
            Distributed
        }

        public enum Vertical
        {
            Top,
            Center,
            Bottom,
            Justify,
            Distributed
        }

        public enum ReadingOrder
        {
            ContextDependent = 0,
            LeftToRight = 1,
            RightToLeft = 2
        }

        public XlsxAlignment(Horizontal horizontal = Horizontal.General, Vertical vertical = Vertical.Bottom,
            int indent = 0, bool justifyLastLine = false, ReadingOrder readingOrder = ReadingOrder.ContextDependent,
            bool shrinkToFit = false, int textRotation = 0, bool wrapText = false)
        {
            HorizontalType = horizontal;
            VerticalType = vertical;
            Indent = indent;
            JustifyLastLine = justifyLastLine;
            ReadingOrderType = readingOrder;
            ShrinkToFit = shrinkToFit;
            TextRotation = textRotation;
            WrapText = wrapText;
        }

        public bool Equals(XlsxAlignment other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return HorizontalType == other.HorizontalType && VerticalType == other.VerticalType
                && Indent == other.Indent && JustifyLastLine == other.JustifyLastLine
                && ReadingOrderType == other.ReadingOrderType && ShrinkToFit == other.ShrinkToFit
                && TextRotation == other.TextRotation && WrapText == other.WrapText;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((XlsxAlignment)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (int)HorizontalType;
                hashCode = (hashCode * 397) ^ (int)VerticalType;
                hashCode = (hashCode * 397) ^ Indent;
                hashCode = (hashCode * 397) ^ JustifyLastLine.GetHashCode();
                hashCode = (hashCode * 397) ^ (int)ReadingOrderType;
                hashCode = (hashCode * 397) ^ ShrinkToFit.GetHashCode();
                hashCode = (hashCode * 397) ^ TextRotation;
                hashCode = (hashCode * 397) ^ WrapText.GetHashCode();
                return hashCode;
            }
        }

        public static bool operator ==(XlsxAlignment left, XlsxAlignment right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(XlsxAlignment left, XlsxAlignment right)
        {
            return !Equals(left, right);
        }
    }
}