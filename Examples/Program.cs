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
using System.Reflection;
using System.Threading.Tasks;

namespace Examples;

public static class Program
{
    public static async Task Main(string[] _)
    {
        Console.WriteLine(System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription);
        Run(Simple.Run);
        Run(MultipleSheet.Run);
        Run(FrozenPanes.Run);
        Run(HideGridlines.Run);
        Run(WorksheetVisibility.Run);
        Run(NumberFormats.Run);
        Run(ColumnFormatting.Run);
        Run(RowFormatting.Run);
        Run(Alignment.Run);
        Run(Border.Run);
        Run(DataValidation.Run);
        Run(RightToLeft.Run);
        Run(SheetProtection.Run);
        Run(HeaderFooterPageBreaks.Run);
        Run(InvalidXmlChars.Run);
        Run(InlineStrings.Run);
        Run(SharedStrings.Run);
        Run(Large.Run);
        await RunAsync(LargeAsync.Run);
        Run(StyledLarge.Run);
        Run(StyledLargeCreateStyles.Run);
        Run(Zip64Huge.Run);
    }

    private static void Run(Action example)
    {
        example();
        WriteMemoryInfo(example.Method);
    }

    private static async Task RunAsync(Func<Task> example)
    {
        await example();
        WriteMemoryInfo(example.Method);
    }

    private static void WriteMemoryInfo(MethodInfo methodInfo)
    {
#if NETCOREAPP3_0_OR_GREATER
        Console.WriteLine($"{methodInfo.DeclaringType!.Name,20}\tGen0: {GC.CollectionCount(0)}\tGen1: {GC.CollectionCount(1)}\tGen2: {GC.CollectionCount(2)}\tTotalMemory: {GC.GetTotalMemory(false)}\tTotalAllocatedBytes: {GC.GetTotalAllocatedBytes()}");
#else
        Console.WriteLine($"{methodInfo.DeclaringType!.Name,20}\tGen0: {GC.CollectionCount(0)}\tGen1: {GC.CollectionCount(1)}\tGen2: {GC.CollectionCount(2)}\tTotalMemory: {GC.GetTotalMemory(false)}");
#endif
    }
}