/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2023 Salvatore ISAJA. All rights reserved.

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

using System.Drawing;
using System.IO;
using LargeXlsx;

namespace Examples;

public static class DataValidation
{
    public static void Run()
    {
        using var stream = new FileStream($"{nameof(DataValidation)}.xlsx", FileMode.Create, FileAccess.Write);
        using var xlsxWriter = new XlsxWriter(stream);
        xlsxWriter.BeginWorksheet("Sheet 1")
            .BeginRow()
            .AddDataValidation(new XlsxDataValidation(validationType: XlsxDataValidation.ValidationType.List, formula1: "\"Lorem,Ipsum,Dolor\""))
            .Write(XlsxStyle.Default.With(new XlsxFill(Color.OldLace)))
            .SkipColumns(1)
            .AddDataValidation(1, 2, new XlsxDataValidation(validationType: XlsxDataValidation.ValidationType.List, formula1: "\"Lorem,Ipsum,Dolor\""))
            .Write(XlsxStyle.Default.With(new XlsxFill(Color.OldLace)), repeatCount: 2)
            .BeginRow()
            .AddDataValidation(XlsxDataValidation.List(new[] { "Lorem", "Ipsum", "Dolor" }))
            .Write(XlsxStyle.Default.With(new XlsxFill(Color.OldLace)))
            .BeginRow()
            .AddDataValidation(XlsxDataValidation.List(new[] { "1", "3.141592", "Sit" },
                showErrorMessage: true, showInputMessage: true, errorStyle: XlsxDataValidation.ErrorStyle.Warning))
            .Write(XlsxStyle.Default.With(new XlsxFill(Color.AliceBlue)))
            .BeginRow()
            .AddDataValidation(XlsxDataValidation.List(new[] { "1", "\"2\"", "\"3.141592\"", "\"3,141592\"", "3,,141592", "\"Sit, Amet\"" },
                showErrorMessage: true, showInputMessage: true, errorTitle: "Error title", error: "A very informative error message",
                promptTitle: "Prompt title", prompt: "A very enlightening prompt"))
            .Write(XlsxStyle.Default.With(new XlsxFill(Color.BlueViolet)))
            .BeginRow()
            .AddDataValidation(new XlsxDataValidation(validationType: XlsxDataValidation.ValidationType.List, formula1: "=Choices!A1:A3"))
            .Write(XlsxStyle.Default.With(new XlsxFill(Color.PaleGreen)))
            .BeginWorksheet("Choices").BeginRow().Write(3.141592).BeginRow().Write("Lorem").BeginRow().Write("Ipsum, Dolor");
    }
}