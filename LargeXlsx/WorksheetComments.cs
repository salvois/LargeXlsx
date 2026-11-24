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
using System.Collections.Generic;

namespace LargeXlsx;

/*
 * Comments cannot be streamed: they live in their own parts (comments and VML drawing),
 * that can only be written after the worksheet part has been closed.
 * Thus they are kept in memory until the end of the worksheet, with memory proportional
 * to the number of comments only.
 */
internal sealed class WorksheetComments
{
    private readonly List<(int Row, int Column, XlsxComment Comment)> _comments = new();
    private readonly HashSet<(int Row, int Column)> _commentedCells = new();

    public int Count => _comments.Count;

    public void Add(int row, int column, XlsxComment comment)
    {
        if (comment == null)
            throw new ArgumentNullException(nameof(comment));
        if (row < 1 || row > Limits.MaxRowCount)
            throw new ArgumentOutOfRangeException(nameof(row), $"Row number must be between 1 and {Limits.MaxRowCount} ({row} specified)");
        if (column < 1 || column > Limits.MaxColumnCount)
            throw new ArgumentOutOfRangeException(nameof(column), $"Column number must be between 1 and {Limits.MaxColumnCount} ({column} specified)");
        // Excel considers a file with more than one comment on the same cell as corrupted
        if (!_commentedCells.Add((row, column)))
            throw new InvalidOperationException($"Cell {Util.GetColumnName(column)}{row} already has a comment");
        _comments.Add((row, column, comment));
    }
}
