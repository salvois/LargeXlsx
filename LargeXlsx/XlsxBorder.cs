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
using System;
using System.Drawing;

namespace LargeXlsx
{
    public class XlsxBorder : IEquatable<XlsxBorder>
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

        public class Line : IEquatable<Line>
        {
            public Color Color { get; }
            public Style Style { get; }

            public Line(Color color, Style style)
            {
                Color = color;
                Style = style;
            }

            #region Equality members
            public bool Equals(Line other)
            {
                if (ReferenceEquals(null, other)) return false;
                if (ReferenceEquals(this, other)) return true;
                return Color.Equals(other.Color) && Style == other.Style;
            }

            public override bool Equals(object obj)
            {
                if (ReferenceEquals(null, obj)) return false;
                if (ReferenceEquals(this, obj)) return true;
                if (obj.GetType() != this.GetType()) return false;
                return Equals((Line)obj);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    return (Color.GetHashCode() * 397) ^ (int)Style;
                }
            }

            public static bool operator ==(Line left, Line right)
            {
                return Equals(left, right);
            }

            public static bool operator !=(Line left, Line right)
            {
                return !Equals(left, right);
            }
            #endregion
        }

        public static readonly XlsxBorder None = new XlsxBorder();

        public Line Top { get; }
        public Line Right { get; }
        public Line Bottom { get; }
        public Line Left { get; }
        public Line Diagonal { get; }
        public bool DiagonalDown { get; }
        public bool DiagonalUp { get; }

        public XlsxBorder(Line top = null, Line right = null, Line bottom = null, Line left = null,
            Line diagonal = null, bool diagonalDown = false, bool diagonalUp = false)
        {
            Top = top;
            Right = right;
            Bottom = bottom;
            Left = left;
            Diagonal = diagonal;
            DiagonalDown = diagonalDown;
            DiagonalUp = diagonalUp;
        }

        public static XlsxBorder Around(Line around) => new XlsxBorder(around, around, around, around);

        #region Equality members
        public bool Equals(XlsxBorder other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Equals(Top, other.Top) && Equals(Right, other.Right) && Equals(Bottom, other.Bottom) && Equals(Left, other.Left)
                   && Equals(Diagonal, other.Diagonal) && DiagonalDown == other.DiagonalDown && DiagonalUp == other.DiagonalUp;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((XlsxBorder)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (Top != null ? Top.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Right != null ? Right.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Bottom != null ? Bottom.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Left != null ? Left.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Diagonal != null ? Diagonal.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ DiagonalDown.GetHashCode();
                hashCode = (hashCode * 397) ^ DiagonalUp.GetHashCode();
                return hashCode;
            }
        }

        public static bool operator ==(XlsxBorder left, XlsxBorder right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(XlsxBorder left, XlsxBorder right)
        {
            return !Equals(left, right);
        }
        #endregion
    }
}