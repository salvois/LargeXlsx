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
namespace LargeXlsx;

/*
 * VML shape ids must be unique across the whole workbook. They are grouped in blocks of 1024 ids,
 * each VML drawing declaring the blocks it uses in <o:idmap>, and Excel uses ids 1 to 1023 of each block.
 * Blocks are allocated when a worksheet is closed, as only then its number of comments is known.
 */
internal sealed class VmlShapeIdAllocator
{
    public const int ShapesPerBlock = 1023;
    private int _nextBlock = 1;

    /// <summary>Reserves enough consecutive blocks for the specified number of shapes.</summary>
    /// <returns>The first reserved block.</returns>
    public int AllocateBlocks(int shapeCount)
    {
        var firstBlock = _nextBlock;
        _nextBlock += (shapeCount + ShapesPerBlock - 1) / ShapesPerBlock;
        return firstBlock;
    }
}
