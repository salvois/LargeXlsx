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
namespace LargeXlsx
{
    public class XlsxHeaderFooter
    {
        public string OddHeader { get; }
        public string OddFooter { get; }
        public string EvenHeader { get; }
        public string EvenFooter { get; }
        public string FirstHeader { get; }
        public string FirstFooter { get; }
        public bool AlignWithMargins { get; }
        public bool ScaleWithDoc { get; }

        public XlsxHeaderFooter(
            string oddHeader = null,
            string oddFooter = null,
            string evenHeader = null,
            string evenFooter = null,
            string firstHeader = null,
            string firstFooter = null,
            bool alignWithMargins = true,
            bool scaleWithDoc = true)
        {
            OddHeader = oddHeader;
            OddFooter = oddFooter;
            EvenHeader = evenHeader;
            EvenFooter = evenFooter;
            FirstHeader = firstHeader;
            FirstFooter = firstFooter;
            AlignWithMargins = alignWithMargins;
            ScaleWithDoc = scaleWithDoc;
        }
    }
}