/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2024 Salvatore ISAJA. All rights reserved.

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
using System.Diagnostics;
using System.IO;
using LargeXlsx;

namespace Examples;

public static class SharedStrings
{
    private const int RowCount = 1_000_000;

    public static void Run()
    {
        var stopwatch = Stopwatch.StartNew();
        DoRun();
        stopwatch.Stop();
        Console.WriteLine($"{nameof(SharedStrings)} completed in {stopwatch.ElapsedMilliseconds} ms.");
    }

    private static void DoRun()
    {
        using var stream = new FileStream($"{nameof(SharedStrings)}.xlsx", FileMode.Create, FileAccess.Write);
        using var xlsxWriter = new XlsxWriter(stream);
        xlsxWriter.BeginWorksheet("Sheet1").BeginRow().Write("This string is not shared");
        for (var i = 0; i < RowCount; i++)
        {
            xlsxWriter.BeginRow()
                .WriteSharedString("Lorem ipsum dolor sit amet")
                .WriteSharedString("consectetur adipiscing elit")
                .WriteSharedString("sed do eiusmod tempor incididunt ut labore et dolore magna aliqua");
        }
        xlsxWriter.BeginWorksheet("Sheet2").BeginRow().Write("This string is not shared either");
        for (var i = 0; i < RowCount; i++)
        {
            xlsxWriter.BeginRow()
                .WriteSharedString("Lorem ipsum dolor sit amet")
                .WriteSharedString("consectetur adipiscing elit")
                .WriteSharedString("sed do eiusmod tempor incididunt ut labore et dolore magna aliqua");
        }
    }
}