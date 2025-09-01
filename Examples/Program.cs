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

namespace Examples;

public static class Program
{
    public static void Main(string[] _)
    {
        var examples = new[]
        {
            Simple.Run,
            MultipleSheet.Run,
            FrozenPanes.Run,
            HideGridlines.Run,
            WorksheetVisibility.Run,
            NumberFormats.Run,
            ColumnFormatting.Run,
            RowFormatting.Run,
            Alignment.Run,
            Border.Run,
            DataValidation.Run,
            RightToLeft.Run,
            Zip64Small.Run,
            SheetProtection.Run,
            HeaderFooterPageBreaks.Run,
            InvalidXmlChars.Run,
            InlineStrings.Run,
            SharedStrings.Run,
            Large.Run,
            StyledLarge.Run,
            StyledLargeCreateStyles.Run,
            //Zip64Huge.Run,
        };
        foreach (var example in examples)
        {
            example();
#if NETCOREAPP3_0_OR_GREATER
            Console.WriteLine($"Gen0: {GC.CollectionCount(0)} Gen1: {GC.CollectionCount(1)} Gen2: {GC.CollectionCount(2)} TotalMemory: {GC.GetTotalMemory(false)} TotalAllocatedBytes: {GC.GetTotalAllocatedBytes()}");
#else
            Console.WriteLine($"Gen0: {GC.CollectionCount(0)} Gen1: {GC.CollectionCount(1)} Gen2: {GC.CollectionCount(2)} TotalMemory: {GC.GetTotalMemory(false)}");
#endif
        }
    }
}