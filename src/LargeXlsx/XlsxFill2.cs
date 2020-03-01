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
    public class XlsxFill2
    {
        public enum Pattern
        {
            None,
            Gray125,
            Solid
        }

        public static readonly XlsxFill2 None = new XlsxFill2(0, Pattern.None, "ffffff");
        public static readonly XlsxFill2 Gray125 = new XlsxFill2(1, Pattern.Gray125, "ffffff");
        internal const int FirstAvailableId = 2; // ids less than 2 are hardcoded by Excel for default fills

        public int Id { get; }
        public Pattern PatternType { get; }
        public string HexRgbColor { get; }

        internal XlsxFill2(int id, Pattern patternType, string hexRgbColor)
        {
            Id = id;
            PatternType = patternType;
            HexRgbColor = hexRgbColor;
        }

        internal static string GetPatternAttributeValue(Pattern patternType)
        {
            switch (patternType)
            {
                case Pattern.None: return "none";
                case Pattern.Gray125: return "gray125";
                case Pattern.Solid: return "solid";
                default: throw new ArgumentOutOfRangeException();
            }
        }
    }
}