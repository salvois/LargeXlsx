/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2026 Salvatore ISAJA. All rights reserved.

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

public static class Comments
{
    public static void Run()
    {
        using var stream = new FileStream($"{nameof(Comments)}.xlsx", FileMode.Create, FileAccess.Write);
        using var xlsxWriter = new XlsxWriter(stream);
        var headerStyle = XlsxStyle.Default.With(new XlsxFill(Color.AliceBlue));
        xlsxWriter.BeginWorksheet("Sheet 1")
            .BeginRow()
            .AddComment(new XlsxComment("This comment is always shown", author: "Alice", visible: true))
            .Write("Product", headerStyle)
            .AddComment(new XlsxComment("Prices include VAT", author: "Bob"))
            .Write("Price", headerStyle)
            .BeginRow()
            .Write("Lorem")
            .AddComment(new XlsxComment("A bigger box\nspanning\nseveral lines", author: "Alice", widthInColumns: 4, heightInRows: 8))
            .Write(3.14)
            .BeginRow()
            .Write("Ipsum")
            .Write(2.72)
            .AddComment(3, 1, new XlsxComment("Comment added after writing the cell, with no author"));
    }
}
