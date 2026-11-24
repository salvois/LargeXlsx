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
using System;

namespace LargeXlsx;

/// <summary>
/// A comment (a.k.a. note) attached to a cell, shown in a pop-up box by spreadsheet applications.
/// </summary>
public class XlsxComment
{
    public string Text { get; }
    public string Author { get; }
    public bool Visible { get; }
    public int WidthInColumns { get; }
    public int HeightInRows { get; }

    /// <param name="text">The plain text of the comment, possibly containing line breaks.</param>
    /// <param name="author">The optional author of the comment.</param>
    /// <param name="visible">Whether the comment box is always shown rather than only when hovering the cell.</param>
    /// <param name="widthInColumns">The width of the comment box, expressed as a number of columns.</param>
    /// <param name="heightInRows">The height of the comment box, expressed as a number of rows.</param>
    public XlsxComment(string text, string author = null, bool visible = false, int widthInColumns = 2, int heightInRows = 4)
    {
        if (widthInColumns < 1)
            throw new ArgumentOutOfRangeException(nameof(widthInColumns), "The width of a comment must be at least one column");
        if (heightInRows < 1)
            throw new ArgumentOutOfRangeException(nameof(heightInRows), "The height of a comment must be at least one row");
        Text = text ?? throw new ArgumentNullException(nameof(text));
        Author = author;
        Visible = visible;
        WidthInColumns = widthInColumns;
        HeightInRows = heightInRows;
    }
}
