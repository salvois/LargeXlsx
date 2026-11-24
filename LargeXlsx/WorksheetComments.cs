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
using System.IO;
using System.Linq;
using System.Threading.Tasks;

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

    /// <summary>
    /// Writes the comments part, the legacy VML drawing part (required by Excel to display comments)
    /// and the worksheet relationships part. Must be called after the worksheet part has been closed,
    /// as zip writers support only one open entry at a time.
    /// </summary>
    public async Task WritePartsAsync(IZipWriter zipWriter, CustomWriter customWriter, int sheetId, int firstVmlBlock, bool skipInvalidCharacters)
    {
        // Explicit positions may be added in any order, but Excel expects comments sorted by cell
        var sortedComments = _comments.OrderBy(c => c.Row).ThenBy(c => c.Column).ToList();
        await WriteEntryAsync(zipWriter, $"xl/comments{sheetId}.xml", stream => WriteCommentsAsync(customWriter, stream, sortedComments, skipInvalidCharacters)).ConfigureAwait(false);
        await WriteEntryAsync(zipWriter, $"xl/drawings/vmlDrawing{sheetId}.vml", stream => WriteVmlDrawingAsync(customWriter, stream, sortedComments, firstVmlBlock)).ConfigureAwait(false);
        await WriteEntryAsync(zipWriter, $"xl/worksheets/_rels/sheet{sheetId}.xml.rels", stream => WriteRelationshipsAsync(customWriter, stream, sheetId)).ConfigureAwait(false);
        // Closed worksheets are kept alive by XlsxWriter until the workbook is disposed,
        // thus release comments as soon as they are written to keep memory bounded to a single worksheet
        _comments.Clear();
        _comments.TrimExcess();
        _commentedCells.Clear();
        _commentedCells.TrimExcess();
    }

    /// <summary>
    /// Computes the value of the x:Anchor element of a VML note, which is what Excel actually uses to place and size
    /// the comment box. It is made of 8 zero-based values: left column, left offset, top row, top offset,
    /// right column, right offset, bottom row, bottom offset, where offsets are in pixels from the top-left
    /// corner of the respective cell. Offsets are fixed to Excel's defaults.
    /// </summary>
    internal static string ComputeAnchor(int row0, int column0, int widthInColumns, int heightInRows)
    {
        // Same default placement as Excel: one column to the right and one row above the commented cell
        // and shifted back when it would go past the last column or row
        var left = Math.Min(column0 + 1, Math.Max(Limits.MaxColumnCount - 1 - widthInColumns, 0));
        var top = Math.Min(Math.Max(row0 - 1, 0), Math.Max(Limits.MaxRowCount - 1 - heightInRows, 0));
        var right = Math.Min(left + widthInColumns, Limits.MaxColumnCount - 1);
        var bottom = Math.Min(top + heightInRows, Limits.MaxRowCount - 1);
        return $"{left}, 15, {top}, 10, {right}, 15, {bottom}, 4";
    }

    private static async Task WriteEntryAsync(IZipWriter zipWriter, string path, Func<Stream, Task> write)
    {
#if NETCOREAPP2_1_OR_GREATER
        await using var stream = zipWriter.CreateEntry(path);
#else
        using var stream = zipWriter.CreateEntry(path);
#endif
        await write(stream).ConfigureAwait(false);
    }

    /*
     * Writes the comments part: the list of distinct authors, then one comment per cell, referencing its author
     * by zero-based index in that list. A null author is written as an empty author, as authorId is mandatory.
     */
    private static async Task WriteCommentsAsync(CustomWriter customWriter, Stream stream, List<(int Row, int Column, XlsxComment Comment)> comments, bool skipInvalidCharacters)
    {
        var authorIds = new Dictionary<string, int>();
        foreach (var c in comments)
        {
            var author = c.Comment.Author ?? "";
            if (!authorIds.ContainsKey(author))
                authorIds.Add(author, authorIds.Count);
        }

        customWriter.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"u8
                            + "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors>\n"u8);
        foreach (var author in authorIds.OrderBy(a => a.Value))
            await customWriter.Append("<author>"u8).AppendEscapedXmlString(author.Key, skipInvalidCharacters).Append("</author>\n"u8).TryFlushToAsync(stream).ConfigureAwait(false);
        customWriter.Append("</authors><commentList>\n"u8);
        foreach (var c in comments)
        {
            await customWriter
                .Append("<comment ref=\""u8).Append(Util.GetUtf8ColumnName(c.Column)).Append(c.Row)
                .Append("\" authorId=\""u8).Append(authorIds[c.Comment.Author ?? ""])
                .Append("\"><text><t"u8).AddSpacePreserveIfNeeded(c.Comment.Text).Append(">"u8)
                .AppendEscapedXmlString(c.Comment.Text, skipInvalidCharacters)
                .Append("</t></text></comment>\n"u8)
                .TryFlushToAsync(stream).ConfigureAwait(false);
        }
        await customWriter.Append("</commentList></comments>\n"u8).FlushToAsync(stream).ConfigureAwait(false);
    }

    /*
     * Writes the legacy VML drawing part, containing one text box shape per comment.
     * The comments part only holds text and authors: Excel needs a VML shape for each comment to display it,
     * and silently drops comments without one. Most of the markup below mirrors what Excel itself writes for
     * a new note, and is required boilerplate rather than something this library makes configurable.
     *
     * Overall structure:
     *   <xml>                    the root of a VML drawing part is literally named "xml"
     *     <o:shapelayout>        layout information for the whole drawing, here only the shape id blocks (o:idmap)
     *     <v:shapetype>          a reusable shape template, here the standard text box, referenced by each shape
     *     <v:shape>*             one shape per comment: look of the box, plus Excel-specific data (x:ClientData)
     *   </xml>
     */
    private static async Task WriteVmlDrawingAsync(CustomWriter customWriter, Stream stream, List<(int Row, int Column, XlsxComment Comment)> comments, int firstVmlBlock)
    {
        // Namespaces: v = VML core (shapes), o = Office extensions (shape layout, shape type ids), x = Excel extensions (client data)
        customWriter.Append("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\n"u8
                            + "<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\""u8);

        // Shape ids must be unique across the whole workbook, and are handed out in blocks of 1024 ids:
        // block n owns ids n * 1024 to n * 1024 + 1023, of which Excel uses the last 1023 (n * 1024 itself is never used).
        // o:idmap must list, comma separated, every block this drawing takes its shape ids from.
        // blockCount is the number of blocks needed for all comments of this worksheet, that is
        // Count / ShapesPerBlock rounded up, using integer arithmetic: 1 to 1023 comments need 1 block,
        // 1024 to 2046 comments need 2 blocks, 2500 comments need 3 blocks.
        // It must match the number of blocks reserved for this worksheet by VmlShapeIdAllocator.AllocateBlocks,
        // which uses the same formula and returned firstVmlBlock, so that the blocks of different worksheets never overlap.
        var blockCount = (comments.Count + VmlShapeIdAllocator.ShapesPerBlock - 1) / VmlShapeIdAllocator.ShapesPerBlock;
        for (var i = 0; i < blockCount; i++)
        {
            if (i > 0) customWriter.Append(","u8);
            customWriter.Append(firstVmlBlock + i);
        }

        // Shape type 202 is the predefined Office text box. The "_x0000_" prefix of ids is a convention of Office VML,
        // with "t" for shape types and "s" for shapes. coordsize is the usual VML coordinate space of 21600 units,
        // in which path draws the rectangular outline of the box (move, lines, close, end).
        customWriter.Append("\"/></o:shapelayout>\n"u8
                            + "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">"u8
                            + "<v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>\n"u8);

        for (var i = 0; i < comments.Count; i++)
        {
            var c = comments[i];
            // The i-th comment takes the (i % 1023)-th usable id of the (i / 1023)-th block reserved for this worksheet,
            // e.g. with firstVmlBlock = 1: comment 0 gets id 1025 (as in Excel), comment 1022 gets 2047, comment 1023 gets 2049.
            var shapeId = (firstVmlBlock + i / VmlShapeIdAllocator.ShapesPerBlock) * 1024 + 1 + i % VmlShapeIdAllocator.ShapesPerBlock;
            customWriter
                // style: the absolute position and size are Excel's defaults, as Excel actually places and sizes the box
                // from x:Anchor below; z-index stacks boxes in comment order; visibility tells whether the box is always
                // shown or only when hovering the commented cell.
                .Append("<v:shape id=\"_x0000_s"u8).Append(shapeId)
                .Append("\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:"u8).Append(i + 1)
                .Append(c.Comment.Visible ? ";visibility:visible"u8 : ";visibility:hidden"u8)
                // The classic look of notes: light yellow plain fill (same color2, thus no gradient) and a black shadow.
                // The text box is left empty, as Excel takes the text from the comments part.
                .Append("\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\"><v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>"u8
                        + "<v:path o:connecttype=\"none\"/><v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"></div></v:textbox>"u8
                        // Excel-specific data turning the shape into a cell note. MoveWithCells and SizeWithCells are
                        // written as Excel does for a new note.
                        + "<x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>"u8)
                .AppendEscapedXmlString(ComputeAnchor(c.Row - 1, c.Column - 1, c.Comment.WidthInColumns, c.Comment.HeightInRows), false)
                // x:Row and x:Column are the zero-based commented cell, linking the shape to the comment with the same ref
                // in the comments part.
                .Append("</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>"u8).Append(c.Row - 1)
                .Append("</x:Row><x:Column>"u8).Append(c.Column - 1)
                .Append("</x:Column>"u8);
            // An empty x:Visible element marks a note that is always shown
            if (c.Comment.Visible)
                customWriter.Append("<x:Visible/>"u8);
            await customWriter.Append("</x:ClientData></v:shape>\n"u8).TryFlushToAsync(stream).ConfigureAwait(false);
        }
        await customWriter.Append("</xml>\n"u8).FlushToAsync(stream).ConfigureAwait(false);
    }

    /*
     * Writes the relationships of the worksheet, linking it to its comments and VML drawing parts.
     * The worksheet references the VML drawing with <legacyDrawing r:id="RidVml1"/>, while the comments part
     * is found by relationship type only.
     */
    private static Task WriteRelationshipsAsync(CustomWriter customWriter, Stream stream, int sheetId) =>
        customWriter
            .Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"u8
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"u8
                    + "<Relationship Id=\"RidCm1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"/xl/comments"u8)
            .Append(sheetId)
            .Append(".xml\"/><Relationship Id=\"RidVml1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"/xl/drawings/vmlDrawing"u8)
            .Append(sheetId)
            .Append(".vml\"/></Relationships>\n"u8)
            .FlushToAsync(stream);
}
